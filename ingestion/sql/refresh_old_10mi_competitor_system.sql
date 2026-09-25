-- Mirrored copy of the fixed-10mi competitor system's refresh procedures, for
-- git-tracked readability and code review. THE DATABASE COPY IS AUTHORITATIVE --
-- this file is kept in sync by hand whenever the live functions change via a new
-- migration. Query pg_get_functiondef('geo.refresh_branch_competitors_10mi_v2'::regproc)
-- (etc.) any time to confirm what's actually live.
--
-- Applied via Supabase migrations "add_full_refresh_procedures_for_old_10mi_system",
-- "fix_generated_column_in_10mi_swap" (2026-09-24), and
-- "cap_10mi_competitors_to_top50_and_fix_future_ordering" (2026-09-25).
--
-- Fixes a real bug found this session: the previous branch_radius_stats_10mi_v2
-- table collapsed every credit union's branches onto one shared key (it used
-- JOIN_NUMBER alone instead of the full CU_NUMBER*1e10 + JOIN_NUMBER*1e5 + SiteId
-- composite id) -- e.g. all ~305 Navy Federal branches shared one row. This rebuild
-- reuses branches_master_v2's already-correct composite key, so every real branch
-- gets its own row. It was also stale (built once, 2026-04-02, never refreshed
-- since) -- these procedures make it repeatable going forward.
--
-- Same competitor-exclusion rule as the proven tiered system
-- (analytics.rebuild_tiered_radius_batch): a branch's own institution never
-- counts as its own competitor. Same CU deposit handling: a CU's institution-wide
-- deposit total is divided by its own branch count before summing into a radius,
-- so a multi-branch CU isn't overcounted.
--
-- geo.branch_competitors_10mi_v2.distance_miles is GENERATED ALWAYS AS
-- (round(distance_m / 1609.344, 2)) -- swap/insert steps must use an explicit
-- column list excluding it, not SELECT *, or Postgres refuses the insert.
--
-- Capped to the top 50 nearest competitors per branch (2026-09-25, Francisco's
-- call): an uncapped flat 10-mile search in dense metros like Manhattan produced
-- 1,500+ "competitors" per branch -- noise for display, and also what made the
-- full rebuild/swap so heavy (12.6M rows). Confirmed first that
-- analytics.branch_opportunity_base's inverse-density scoring reads from
-- geo.branch_radius_stats_tiered_v1 (the ADAPTIVE tiered system) and never
-- touches these flat-10mi tables at all, so the cap has zero effect on scoring.
-- branch_radius_stats_10mi_v2's own counts/deposits must stay accurate (computed
-- from the FULL candidate set, not the capped one it's rendered from), so the
-- rebuild ordering changed: geo.refresh_branch_competitors_10mi_v2() no longer
-- drops its full (uncapped) staging table (geo.branch_competitors_10mi_v2_new)
-- immediately after capping+swapping the production pairs table -- it leaves that
-- staging table behind for geo.refresh_branch_radius_stats_10mi_v2() to aggregate
-- from, and THAT procedure drops it once done. The two must still run in that
-- order, which public.refresh_old_10mi_competitor_system() already enforces.
--
-- Usage:
--   CALL public.refresh_old_10mi_competitor_system();   -- does both tables, in order
-- or individually (competitors MUST run first):
--   CALL geo.refresh_branch_competitors_10mi_v2();
--   CALL geo.refresh_branch_radius_stats_10mi_v2();
--
-- Batches the branch x branch spatial join in chunks of 5,000 "my" branches with a
-- COMMIT between each (procedure, not function, so this is legal) -- a single
-- unbatched 93k x 93k join hit a statement timeout when first attempted. Backs up
-- both tables to backup.<table>_YYYYMMDD before swapping; aborts with no changes
-- made if the rebuilt row count looks implausible.
--
-- First live run (2026-09-24, 2025->2026 branches_master_v2 refresh) completed by
-- finishing the batch loop by hand after the Supabase Studio SQL editor's own
-- browser-side timeout cut the connection mid-run (query kept running server-side,
-- confirmed via pg_stat_activity, but nothing was left listening for the result).
-- The swap itself also failed once on the generated-column bug (fixed) and once
-- more on Studio's timeout for a 12.6M-row bulk copy (also completed server-side
-- despite the client-side error). Long term this argues for running these from
-- something other than the browser SQL editor (a16 Railway scheduling, or psql).

CREATE OR REPLACE FUNCTION geo.rebuild_10mi_competitors_batch(p_offset integer, p_limit integer)
RETURNS integer LANGUAGE plpgsql AS $$
DECLARE
  v_count integer;
  v_year integer;
BEGIN
  SELECT MAX(year) INTO v_year FROM geo.branches_master_v2;

  WITH my_batch AS (
    SELECT branch_id, bank_id, institution_type, geom
    FROM geo.branches_master_v2
    WHERE year = v_year
    ORDER BY branch_id
    OFFSET p_offset LIMIT p_limit
  )
  INSERT INTO geo.branch_competitors_10mi_v2_new
    (my_branch_id, my_bank_id, my_institution_type, competitor_bank_id, competitor_branch_uninumbr,
     competitor_institution_type, distance_m, radius_m, as_of_date, distance_miles)
  SELECT
    a.branch_id, a.bank_id, a.institution_type,
    b.bank_id, b.branch_id, b.institution_type,
    ST_Distance(a.geom, b.geom),
    16093.4,
    CURRENT_DATE,
    ROUND((ST_Distance(a.geom, b.geom) / 1609.34)::numeric, 3)
  FROM my_batch a
  JOIN geo.branches_master_v2 b
    ON b.year = v_year
    AND b.bank_id <> a.bank_id
    AND ST_DWithin(a.geom, b.geom, 16093.4);

  GET DIAGNOSTICS v_count = ROW_COUNT;
  RETURN v_count;
END;
$$;

-- NOTE: the batch function above still inserts into geo.branch_competitors_10mi_v2_new
-- (LIKE geo.branch_competitors_10mi_v2 INCLUDING DEFAULTS), which does NOT carry
-- generation expressions -- so distance_miles is a plain computed value on that
-- staging table (harmless; only the PRODUCTION table's version is generated).

CREATE OR REPLACE PROCEDURE geo.refresh_branch_competitors_10mi_v2()
LANGUAGE plpgsql AS $$
DECLARE
  v_year integer;
  v_total_branches integer;
  v_batch_size integer := 5000;
  v_offset integer := 0;
  v_full_count bigint;
  v_capped_count bigint;
BEGIN
  SELECT MAX(year) INTO v_year FROM geo.branches_master_v2;
  SELECT COUNT(*) INTO v_total_branches FROM geo.branches_master_v2 WHERE year = v_year;

  DROP TABLE IF EXISTS geo.branch_competitors_10mi_v2_new;
  CREATE TABLE geo.branch_competitors_10mi_v2_new (LIKE geo.branch_competitors_10mi_v2 INCLUDING DEFAULTS);
  COMMIT;

  WHILE v_offset < v_total_branches LOOP
    PERFORM geo.rebuild_10mi_competitors_batch(v_offset, v_batch_size);
    v_offset := v_offset + v_batch_size;
    COMMIT;
  END LOOP;

  CREATE INDEX ON geo.branch_competitors_10mi_v2_new (my_branch_id);
  CREATE INDEX ON geo.branch_competitors_10mi_v2_new (competitor_branch_uninumbr);
  ANALYZE geo.branch_competitors_10mi_v2_new;
  COMMIT;

  SELECT COUNT(*) INTO v_full_count FROM geo.branch_competitors_10mi_v2_new;
  IF v_full_count < v_total_branches * 5 THEN
    RAISE EXCEPTION 'branch_competitors_10mi_v2 rebuild produced only % rows for % branches - aborting before swap', v_full_count, v_total_branches;
  END IF;

  -- Cap to top-50 nearest per branch for what actually gets stored/served.
  -- branch_radius_stats_10mi_v2 (built next, by the OTHER procedure) reads
  -- from the FULL geo.branch_competitors_10mi_v2_new BEFORE it gets dropped
  -- -- so it stays accurate regardless of this cap.
  DROP TABLE IF EXISTS geo.branch_competitors_10mi_v2_capped;
  CREATE TABLE geo.branch_competitors_10mi_v2_capped (LIKE geo.branch_competitors_10mi_v2 INCLUDING DEFAULTS);
  INSERT INTO geo.branch_competitors_10mi_v2_capped
    (my_branch_id, my_bank_id, my_institution_type, competitor_bank_id, competitor_branch_uninumbr,
     competitor_institution_type, distance_m, radius_m, as_of_date)
  SELECT my_branch_id, my_bank_id, my_institution_type, competitor_bank_id, competitor_branch_uninumbr,
         competitor_institution_type, distance_m, radius_m, as_of_date
  FROM (
    SELECT *, ROW_NUMBER() OVER (PARTITION BY my_branch_id ORDER BY distance_m) AS rn
    FROM geo.branch_competitors_10mi_v2_new
  ) t
  WHERE rn <= 50;
  COMMIT;

  SELECT COUNT(*) INTO v_capped_count FROM geo.branch_competitors_10mi_v2_capped;

  EXECUTE format('DROP TABLE IF EXISTS backup.branch_competitors_10mi_v2_%s', to_char(CURRENT_DATE, 'YYYYMMDD'));
  EXECUTE format('CREATE TABLE backup.branch_competitors_10mi_v2_%s AS TABLE geo.branch_competitors_10mi_v2', to_char(CURRENT_DATE, 'YYYYMMDD'));

  TRUNCATE geo.branch_competitors_10mi_v2;
  INSERT INTO geo.branch_competitors_10mi_v2
    (my_branch_id, my_bank_id, my_institution_type, competitor_bank_id, competitor_branch_uninumbr,
     competitor_institution_type, distance_m, radius_m, as_of_date)
  SELECT my_branch_id, my_bank_id, my_institution_type, competitor_bank_id, competitor_branch_uninumbr,
         competitor_institution_type, distance_m, radius_m, as_of_date
  FROM geo.branch_competitors_10mi_v2_capped;

  GRANT SELECT ON geo.branch_competitors_10mi_v2 TO anon, authenticated, service_role;

  DROP TABLE geo.branch_competitors_10mi_v2_capped;
  COMMIT;
  -- geo.branch_competitors_10mi_v2_new (the FULL uncapped set) is deliberately
  -- NOT dropped here -- geo.refresh_branch_radius_stats_10mi_v2() reads from it
  -- next and drops it when done. Must run in that order (the umbrella procedure
  -- below already does).

  RAISE NOTICE 'geo.branch_competitors_10mi_v2 refreshed: % pairs capped to top-50 (% total before cap) for % branches (year %)', v_capped_count, v_full_count, v_total_branches, v_year;
END;
$$;

-- Sources from the FULL uncapped staging table left behind by the competitors
-- procedure above, not from the (now-capped) production pairs table -- keeps
-- branches_in_radius / total_deposits_10mi accurate for dense branches.
CREATE OR REPLACE PROCEDURE geo.refresh_branch_radius_stats_10mi_v2()
LANGUAGE plpgsql AS $$
DECLARE
  v_year integer;
  v_total_branches integer;
  v_inserted bigint;
BEGIN
  SELECT MAX(year) INTO v_year FROM geo.branches_master_v2;
  SELECT COUNT(*) INTO v_total_branches FROM geo.branches_master_v2 WHERE year = v_year;

  IF to_regclass('geo.branch_competitors_10mi_v2_new') IS NULL THEN
    RAISE EXCEPTION 'geo.branch_competitors_10mi_v2_new (full uncapped staging) not found - run geo.refresh_branch_competitors_10mi_v2() first';
  END IF;

  DROP TABLE IF EXISTS geo.branch_radius_stats_10mi_v2_new;
  CREATE TABLE geo.branch_radius_stats_10mi_v2_new (LIKE geo.branch_radius_stats_10mi_v2 INCLUDING DEFAULTS);

  WITH branch_counts AS (
    SELECT bank_id, COUNT(*) AS n FROM geo.branches_master_v2
    WHERE year = v_year GROUP BY bank_id
  ),
  agg AS (
    SELECT
      c.my_branch_id,
      COUNT(*) AS branches_in_radius,
      COUNT(*) FILTER (WHERE c.competitor_institution_type = 'bank') AS bank_branches_in_radius,
      COUNT(*) FILTER (WHERE c.competitor_institution_type = 'cu') AS cu_branches_in_radius,
      SUM(CASE WHEN c.competitor_institution_type = 'cu'
               THEN bm.depsum_branch::numeric / NULLIF(bc.n,0)
               ELSE bm.depsum_branch::numeric END) AS total_deposits_10mi,
      SUM(CASE WHEN c.competitor_institution_type = 'bank' THEN bm.depsum_branch::numeric ELSE 0 END) AS bank_deposits_10mi,
      SUM(CASE WHEN c.competitor_institution_type = 'cu'
               THEN bm.depsum_branch::numeric / NULLIF(bc.n,0) ELSE 0 END) AS cu_deposits_10mi
    FROM geo.branch_competitors_10mi_v2_new c
    JOIN geo.branches_master_v2 bm ON bm.branch_id = c.competitor_branch_uninumbr
    LEFT JOIN branch_counts bc ON bc.bank_id = c.competitor_bank_id
    GROUP BY c.my_branch_id
  )
  INSERT INTO geo.branch_radius_stats_10mi_v2_new
    (my_branch_id, bank_id, institution_type, bank_name, branch_name, state, zip,
     branches_in_radius, bank_branches_in_radius, cu_branches_in_radius,
     total_deposits_10mi, bank_deposits_10mi, cu_deposits_10mi,
     avg_deposits_per_branch_10mi, cu_deposit_share_pct)
  SELECT
    m.branch_id, m.bank_id, m.institution_type, m.bank_name, m.branch_name, m.state, m.zip,
    COALESCE(a.branches_in_radius, 0),
    COALESCE(a.bank_branches_in_radius, 0),
    COALESCE(a.cu_branches_in_radius, 0),
    COALESCE(a.total_deposits_10mi, 0),
    COALESCE(a.bank_deposits_10mi, 0),
    COALESCE(a.cu_deposits_10mi, 0),
    CASE WHEN COALESCE(a.branches_in_radius,0) = 0 THEN NULL
         ELSE ROUND(a.total_deposits_10mi / a.branches_in_radius, 2) END,
    CASE WHEN COALESCE(a.total_deposits_10mi,0) = 0 THEN NULL
         ELSE ROUND((a.cu_deposits_10mi / a.total_deposits_10mi * 100)::numeric, 2) END
  FROM geo.branches_master_v2 m
  LEFT JOIN agg a ON a.my_branch_id = m.branch_id
  WHERE m.year = v_year;

  CREATE UNIQUE INDEX ON geo.branch_radius_stats_10mi_v2_new (my_branch_id);
  ANALYZE geo.branch_radius_stats_10mi_v2_new;

  SELECT COUNT(*) INTO v_inserted FROM geo.branch_radius_stats_10mi_v2_new;
  IF v_inserted < v_total_branches * 0.95 THEN
    RAISE EXCEPTION 'branch_radius_stats_10mi_v2 rebuild produced % rows for % branches - aborting before swap', v_inserted, v_total_branches;
  END IF;

  EXECUTE format('DROP TABLE IF EXISTS backup.branch_radius_stats_10mi_v2_%s', to_char(CURRENT_DATE, 'YYYYMMDD'));
  EXECUTE format('CREATE TABLE backup.branch_radius_stats_10mi_v2_%s AS TABLE geo.branch_radius_stats_10mi_v2', to_char(CURRENT_DATE, 'YYYYMMDD'));

  TRUNCATE geo.branch_radius_stats_10mi_v2;
  INSERT INTO geo.branch_radius_stats_10mi_v2 SELECT * FROM geo.branch_radius_stats_10mi_v2_new;

  GRANT SELECT ON geo.branch_radius_stats_10mi_v2 TO anon, authenticated, service_role;

  DROP TABLE geo.branch_radius_stats_10mi_v2_new;
  DROP TABLE IF EXISTS geo.branch_competitors_10mi_v2_new;

  RAISE NOTICE 'geo.branch_radius_stats_10mi_v2 refreshed: % branches (year %)', v_inserted, v_year;
END;
$$;

CREATE OR REPLACE PROCEDURE public.refresh_old_10mi_competitor_system()
LANGUAGE plpgsql AS $$
BEGIN
  CALL geo.refresh_branch_competitors_10mi_v2();
  CALL geo.refresh_branch_radius_stats_10mi_v2();
END;
$$;
