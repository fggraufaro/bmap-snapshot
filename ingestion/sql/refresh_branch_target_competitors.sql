-- Mirrored copy of the target-competitor engine's refresh procedure, for
-- git-tracked readability and code review. THE DATABASE COPY IS AUTHORITATIVE --
-- this file is kept in sync by hand whenever the live functions change via a
-- new migration. Query pg_get_functiondef('analytics.refresh_branch_target_competitors'::regproc)
-- (etc.) any time to confirm what's actually live.
--
-- Applied via Supabase migration "add_refresh_branch_target_competitors_procedure"
-- (2026-09-25).
--
-- analytics.branch_target_competitors (the Hub's pre-scored target-competitor
-- drill-down engine, "top 3 vulnerable targets per branch") was stale: last
-- rebuilt 2026-09-18 (a55), before branches_master_v2 moved to 2026. Checked
-- directly: ~2.3% of its branch references pointed to branches no longer in
-- the 2026 universe, and ~6% of *current* 2026 branches had no row at all.
--
-- Fixed the same way as the tiered competitor system earlier this week: the
-- existing analytics.rebuild_btc_batch() batched through ref.dim_institutions
-- (a static table, same staleness risk found there before) and joined it
-- twice just to read institution_type -- which analytics.branch_opportunity_base
-- already carries directly. Rewritten to batch off branch_opportunity_base's
-- own distinct inst_keys (always current) and dropped both ref.dim_institutions
-- joins entirely. No new data source needed -- everything it draws from
-- (branch_opportunity_base, branches_master_v2, bank_financial_snapshot_latest)
-- was already rebuilt for 2026 earlier this week. Scoring logic (the
-- vuln_score formula, top-3-per-branch ranking, the 10x-500% deposit-size
-- band filter) is unchanged from the proven a55 version -- only the
-- staleness-causing joins were touched.
--
-- Batches by INSTITUTION, not by branch (unlike the tiered/10mi rebuilds) --
-- batch size 200 institutions worked reliably; institution size varies
-- widely (a single large bank can dominate one batch), so watch for uneven
-- per-batch row counts, not a bug.
--
-- Usage:
--   CALL analytics.refresh_branch_target_competitors();
--
-- First live run (2026-09-25) completed via manual batch-by-batch execution
-- (same as the tiered/10mi systems' first runs) after the wrapping
-- procedure's internal batching loop hit a genuine statement timeout through
-- the Supabase MCP tool's connection -- unclear root cause (not the
-- PostgREST/authenticator 8s issue, since this wasn't called via PostgREST),
-- possibly just batch-size sensitivity given how uneven institution sizes
-- are. Batch size 200 was reliable when run individually. Result:
-- 256,384 rows, 89,985 distinct branches, 100% match with the 2026 branch
-- universe (was 97.7%).

CREATE OR REPLACE FUNCTION analytics.rebuild_btc_batch(p_offset integer, p_limit integer)
RETURNS integer LANGUAGE plpgsql AS $$
DECLARE
  v_count integer;
BEGIN
  WITH batch_insts AS (
    SELECT DISTINCT inst_key FROM analytics.branch_opportunity_base
    ORDER BY inst_key OFFSET p_offset LIMIT p_limit
  ),
  my_branches AS (
    SELECT b.uninumbr, g.lat AS my_lat, g.lon AS my_lon, g.bank_id AS my_bank_id,
           b.latest_dep::numeric AS branch_dep, b.inst_key, b.institution_type AS my_inst_type
    FROM analytics.branch_opportunity_base b
    JOIN geo.branches_master_v2 g ON g.branch_id = b.uninumbr
    WHERE b.inst_key IN (SELECT inst_key FROM batch_insts)
  ),
  bank_totals AS (
    SELECT inst_key, SUM(latest_dep::numeric) AS total
    FROM analytics.branch_opportunity_base
    WHERE inst_key IN (SELECT inst_key FROM batch_insts)
    GROUP BY inst_key
  ),
  latest_year AS (SELECT MAX(year) AS yr FROM geo.branches_master_v2),
  branch_counts AS (
    SELECT bank_id, COUNT(*) AS n FROM geo.branches_master_v2
    WHERE year = (SELECT yr FROM latest_year) GROUP BY bank_id
  ),
  candidates AS (
    SELECT mb.uninumbr AS my_branch_id, gm.branch_id AS target_uninumbr, gm.bank_id,
      ST_Distance(gm.geom, ST_SetSRID(ST_MakePoint(mb.my_lon, mb.my_lat),4326)::geography) / 1609.34 AS distance_miles
    FROM my_branches mb
    JOIN geo.branches_master_v2 gm ON gm.year = (SELECT yr FROM latest_year)
    JOIN branch_counts bc ON bc.bank_id = gm.bank_id
    WHERE gm.bank_id != mb.my_bank_id
      AND ST_DWithin(gm.geom, ST_SetSRID(ST_MakePoint(mb.my_lon, mb.my_lat),4326)::geography, 10*1609.34)
  ),
  density AS (
    SELECT my_branch_id, COUNT(*) FILTER (WHERE distance_miles <= 1.0) AS density_1mi
    FROM candidates GROUP BY my_branch_id
  ),
  dep_1mi AS (
    SELECT c.my_branch_id,
      COALESCE(SUM(CASE WHEN c.distance_miles <= 1.0 THEN
        (CASE WHEN gm.institution_type='cu' THEN gm.depsum_branch::numeric/NULLIF(bc.n,0) ELSE gm.depsum_branch::numeric END)
      ELSE 0 END),0) AS deposits_1mi
    FROM candidates c
    JOIN geo.branches_master_v2 gm ON gm.branch_id = c.target_uninumbr AND gm.year=(SELECT yr FROM latest_year)
    JOIN branch_counts bc ON bc.bank_id = gm.bank_id
    GROUP BY c.my_branch_id
  ),
  radius_calc AS (
    SELECT mb.uninumbr, mb.branch_dep, mb.my_inst_type, mb.inst_key,
      COALESCE(d.density_1mi,0) AS density_1mi,
      COALESCE(dm.deposits_1mi,0) AS deposits_1mi,
      bt.total AS bank_total,
      analytics.determine_adaptive_radius(COALESCE(d.density_1mi,0)::int, mb.branch_dep, COALESCE(dm.deposits_1mi,0), bt.total) AS radius
    FROM my_branches mb
    LEFT JOIN density d ON d.my_branch_id = mb.uninumbr
    LEFT JOIN dep_1mi dm ON dm.my_branch_id = mb.uninumbr
    JOIN bank_totals bt ON bt.inst_key = mb.inst_key
  ),
  filtered AS (
    SELECT c.my_branch_id, c.target_uninumbr, r.inst_key AS my_inst_key, r.my_inst_type
    FROM candidates c
    JOIN radius_calc r ON r.uninumbr = c.my_branch_id
    JOIN analytics.branch_opportunity_base ob2 ON ob2.uninumbr = c.target_uninumbr
    WHERE c.distance_miles <= r.radius
      AND ob2.latest_dep::numeric BETWEEN r.branch_dep * 0.10 AND r.branch_dep * 5.0
  ),
  scored AS (
    SELECT
      c.my_branch_id, f.my_inst_key,
      ob.uninumbr AS target_uninumbr, ob.inst_key AS target_inst_key,
      ob.namefull AS target_namefull, ob.namebr AS target_namebr,
      ob.institution_type AS target_institution_type,
      ROUND(ob.latest_dep::numeric / 1e6, 1) AS target_dep_M,
      ROUND(ob.yoy_deposits::numeric * 100, 2) AS target_yoy_pct,
      ROUND(ob.opportunity_score::numeric, 1) AS target_opp_score,
      ob.opportunity_zone AS target_zone,
      ROUND(c.distance_miles::numeric, 2) AS target_dist_mi,
      fin.roa AS target_roa, fin.noncurrent_assets_pct AS target_noncurrent_pct,
      ROUND((
          (100 - ob.opportunity_score::numeric)
          * CASE WHEN ob.yoy_deposits::numeric < 0 THEN 1.5 ELSE 1 END
          * CASE WHEN ob.institution_type = f.my_inst_type THEN 1.3 ELSE 0.7 END
          * CASE WHEN COALESCE(fin.roa::numeric, 1) < 0.005 THEN 1.4 ELSE 1 END
          * CASE WHEN COALESCE(fin.noncurrent_assets_pct::numeric, 0) > 0.02 THEN 1.3 ELSE 1 END
          * CASE WHEN c.distance_miles::numeric <= 1 THEN 1.3 ELSE 1 END
      )::numeric, 2) AS vuln_score
    FROM filtered f
    JOIN candidates c ON c.my_branch_id = f.my_branch_id AND c.target_uninumbr = f.target_uninumbr
    JOIN analytics.branch_opportunity_base ob ON ob.uninumbr = f.target_uninumbr
    LEFT JOIN analytics.bank_financial_snapshot_latest fin ON fin.inst_key = ob.inst_key
  ),
  ranked AS (
    SELECT *, ROW_NUMBER() OVER (PARTITION BY my_branch_id ORDER BY vuln_score DESC) AS target_rank
    FROM scored
  )
  INSERT INTO analytics.branch_target_competitors_new
    (my_branch_id, my_inst_key, target_uninumbr, target_inst_key, target_namefull, target_namebr,
     target_institution_type, target_dep_m, target_yoy_pct, target_opp_score, target_zone,
     target_dist_mi, target_roa, target_noncurrent_pct, vuln_score, target_rank)
  SELECT my_branch_id, my_inst_key, target_uninumbr, target_inst_key, target_namefull, target_namebr,
    target_institution_type, target_dep_M, target_yoy_pct, target_opp_score, target_zone,
    target_dist_mi, target_roa, target_noncurrent_pct, vuln_score, target_rank
  FROM ranked WHERE target_rank <= 3;

  GET DIAGNOSTICS v_count = ROW_COUNT;
  RETURN v_count;
END;
$$;

CREATE OR REPLACE PROCEDURE analytics.refresh_branch_target_competitors()
LANGUAGE plpgsql AS $$
DECLARE
  v_total_insts integer;
  v_batch_size integer := 500;
  v_offset integer := 0;
  v_inserted bigint;
  v_current_count bigint;
BEGIN
  SET statement_timeout = '30min';

  SELECT COUNT(*) INTO v_current_count FROM analytics.branch_target_competitors;
  SELECT COUNT(DISTINCT inst_key) INTO v_total_insts FROM analytics.branch_opportunity_base;

  DROP TABLE IF EXISTS analytics.branch_target_competitors_new;
  CREATE TABLE analytics.branch_target_competitors_new (LIKE analytics.branch_target_competitors INCLUDING DEFAULTS);
  COMMIT;
  SET statement_timeout = '30min';

  WHILE v_offset < v_total_insts LOOP
    PERFORM analytics.rebuild_btc_batch(v_offset, v_batch_size);
    v_offset := v_offset + v_batch_size;
    COMMIT;
    SET statement_timeout = '30min';
  END LOOP;

  CREATE INDEX ON analytics.branch_target_competitors_new (my_branch_id);
  ANALYZE analytics.branch_target_competitors_new;
  COMMIT;
  SET statement_timeout = '30min';

  SELECT COUNT(*) INTO v_inserted FROM analytics.branch_target_competitors_new;
  IF v_current_count > 0 AND (v_inserted < v_current_count * 0.5 OR v_inserted > v_current_count * 2.0) THEN
    RAISE EXCEPTION 'branch_target_competitors rebuild produced % rows vs current % - outside expected range, aborting before swap', v_inserted, v_current_count;
  END IF;

  EXECUTE format('DROP TABLE IF EXISTS backup.branch_target_competitors_%s', to_char(CURRENT_DATE, 'YYYYMMDD'));
  EXECUTE format('CREATE TABLE backup.branch_target_competitors_%s AS TABLE analytics.branch_target_competitors', to_char(CURRENT_DATE, 'YYYYMMDD'));

  TRUNCATE analytics.branch_target_competitors;
  INSERT INTO analytics.branch_target_competitors
    (my_branch_id, my_inst_key, target_uninumbr, target_inst_key, target_namefull, target_namebr,
     target_institution_type, target_dep_m, target_yoy_pct, target_opp_score, target_zone,
     target_dist_mi, target_roa, target_noncurrent_pct, vuln_score, target_rank)
  SELECT my_branch_id, my_inst_key, target_uninumbr, target_inst_key, target_namefull, target_namebr,
    target_institution_type, target_dep_m, target_yoy_pct, target_opp_score, target_zone,
    target_dist_mi, target_roa, target_noncurrent_pct, vuln_score, target_rank
  FROM analytics.branch_target_competitors_new;

  GRANT SELECT ON analytics.branch_target_competitors TO anon, authenticated, service_role;

  DROP TABLE analytics.branch_target_competitors_new;
  COMMIT;

  RAISE NOTICE 'branch_target_competitors refreshed: % rows (was %)', v_inserted, v_current_count;
END;
$$;
