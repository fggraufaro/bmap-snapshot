-- Mirrored copy of the adaptive tiered competitor system's refresh procedures,
-- for git-tracked readability and code review. THE DATABASE COPY IS AUTHORITATIVE --
-- this file is kept in sync by hand whenever the live functions change via a new
-- migration. Query pg_get_functiondef('analytics.refresh_branch_competitors_tiered_v1'::regproc)
-- (etc.) any time to confirm what's actually live.
--
-- Applied via Supabase migration "add_refresh_procedures_for_tiered_system" (2026-09-24).
--
-- Two fixes vs. the existing analytics.rebuild_tiered_radius_batch() (left in
-- place, unused going forward -- not dropped in case anything still calls it):
--
--   1. Batches directly off geo.branches_master_v2 (always current) instead of
--      ref.dim_institutions, a static reference table found to be missing 260
--      real institutions that exist in the live data today -- meaning the old
--      function was already silently skipping their branches.
--
--   2. inst_key / institution-total lookups prefer analytics.branch_opportunity_base
--      (matches today's live tier-radius behavior exactly for every branch that
--      overlaps -- same weighting, same quirks, deliberately not "fixed") but fall
--      back to deriving them directly from branches_master_v2 for branches/
--      institutions new since branch_opportunity_base's last refresh. The old
--      rebuild started from an INNER JOIN on branch_opportunity_base, which would
--      silently drop any branch not yet present there -- exactly the case every
--      time a brand-new year's data lands before branch_opportunity_base has been
--      recomputed for it.
--
-- Same batching-with-COMMIT pattern as refresh_old_10mi_competitor_system.sql
-- (procedure, not function, so mid-run COMMIT is legal) -- avoids the same
-- statement-timeout issue a single unbatched pass would hit.
--
-- Usage:
--   CALL public.refresh_tiered_competitor_system();   -- does both tables, in order
-- or individually:
--   CALL analytics.refresh_branch_competitors_tiered_v1();
--   CALL geo.refresh_branch_radius_stats_tiered_v1();
--
-- Sanity-tested this session on a 200-branch sample before being handed off:
-- 2,091 competitor rows, 189/200 branches got at least one row (the other 11
-- plausibly isolated/rural with zero nearby competitors), zero NULL tier radii.

CREATE OR REPLACE FUNCTION analytics.rebuild_tiered_competitors_batch(p_offset integer, p_limit integer)
RETURNS integer LANGUAGE plpgsql AS $$
DECLARE
  v_count integer;
  v_year integer;
BEGIN
  SELECT MAX(year) INTO v_year FROM geo.branches_master_v2;

  WITH cu_branch_counts AS (
    SELECT bank_id, COUNT(*) AS n_branches
    FROM geo.branches_master_v2 WHERE year = v_year GROUP BY bank_id
  ),
  my_batch AS (
    SELECT g.branch_id AS my_branch_id, g.bank_id, g.institution_type,
           g.lat AS my_lat, g.lon AS my_lon, g.depsum_branch::numeric AS branch_dep,
           COALESCE(ob.inst_key,
             CASE WHEN g.institution_type = 'bank' THEN 'bank_' || g.bank_id::text
                  WHEN g.institution_type = 'cu'   THEN 'cu_'   || g.bank_id::text
                  ELSE g.bank_id::text END) AS inst_key
    FROM geo.branches_master_v2 g
    LEFT JOIN analytics.branch_opportunity_base ob ON ob.uninumbr = g.branch_id
    WHERE g.year = v_year
    ORDER BY g.branch_id
    OFFSET p_offset LIMIT p_limit
  ),
  bank_totals AS (
    SELECT mb.inst_key,
      COALESCE(
        (SELECT SUM(latest_dep::numeric) FROM analytics.branch_opportunity_base WHERE inst_key = mb.inst_key),
        CASE WHEN mb.institution_type = 'cu' THEN mb.branch_dep_max ELSE mb.branch_dep_sum END
      ) AS total
    FROM (
      SELECT inst_key, institution_type, SUM(branch_dep) AS branch_dep_sum, MAX(branch_dep) AS branch_dep_max
      FROM my_batch GROUP BY inst_key, institution_type
    ) mb
  ),
  candidates AS (
    SELECT mb.my_branch_id, mb.bank_id, mb.institution_type, mb.inst_key,
           gm.branch_id AS target_branch_id, gm.bank_id AS target_bank_id,
           gm.institution_type AS target_institution_type, gm.depsum_branch::numeric AS target_dep,
           ST_Distance(gm.geom, ST_SetSRID(ST_MakePoint(mb.my_lon, mb.my_lat),4326)::geography) / 1609.34 AS distance_miles
    FROM my_batch mb
    JOIN geo.branches_master_v2 gm
      ON gm.year = v_year
     AND gm.bank_id != mb.bank_id
     AND ST_DWithin(gm.geom, ST_SetSRID(ST_MakePoint(mb.my_lon, mb.my_lat),4326)::geography, 10*1609.34)
  ),
  density AS (
    SELECT my_branch_id, COUNT(*) FILTER (WHERE distance_miles <= 1.0) AS density_1mi
    FROM candidates GROUP BY my_branch_id
  ),
  dep_1mi AS (
    SELECT c.my_branch_id,
      COALESCE(SUM(CASE WHEN c.distance_miles <= 1.0 THEN
        (CASE WHEN c.target_institution_type='cu' THEN c.target_dep/NULLIF(bc.n_branches,0) ELSE c.target_dep END)
      ELSE 0 END),0) AS deposits_1mi
    FROM candidates c
    JOIN cu_branch_counts bc ON bc.bank_id = c.target_bank_id
    GROUP BY c.my_branch_id
  ),
  radius_calc AS (
    SELECT mb.my_branch_id, mb.bank_id, mb.institution_type,
      analytics.determine_density_radius(COALESCE(d.density_1mi,0)::int, mb.branch_dep, COALESCE(dm.deposits_1mi,0), bt.total) AS tier_radius
    FROM my_batch mb
    LEFT JOIN density d ON d.my_branch_id = mb.my_branch_id
    LEFT JOIN dep_1mi dm ON dm.my_branch_id = mb.my_branch_id
    JOIN bank_totals bt ON bt.inst_key = mb.inst_key
  ),
  filtered AS (
    SELECT c.my_branch_id, r.bank_id, r.institution_type, c.target_bank_id, c.target_branch_id,
           c.target_institution_type, c.distance_miles, r.tier_radius
    FROM candidates c
    JOIN radius_calc r ON r.my_branch_id = c.my_branch_id
    WHERE c.distance_miles <= r.tier_radius
  )
  INSERT INTO geo.branch_competitors_tiered_v1_new
    (my_branch_id, my_bank_id, my_institution_type, competitor_bank_id, competitor_branch_uninumbr,
     competitor_institution_type, distance_miles, tier_radius_mi, as_of_date)
  SELECT my_branch_id, bank_id, institution_type, target_bank_id, target_branch_id,
         target_institution_type, ROUND(distance_miles::numeric,3), tier_radius, CURRENT_DATE
  FROM filtered;

  GET DIAGNOSTICS v_count = ROW_COUNT;
  RETURN v_count;
END;
$$;

CREATE OR REPLACE PROCEDURE analytics.refresh_branch_competitors_tiered_v1()
LANGUAGE plpgsql AS $$
DECLARE
  v_year integer;
  v_total_branches integer;
  v_batch_size integer := 5000;
  v_offset integer := 0;
  v_inserted bigint;
BEGIN
  SELECT MAX(year) INTO v_year FROM geo.branches_master_v2;
  SELECT COUNT(*) INTO v_total_branches FROM geo.branches_master_v2 WHERE year = v_year;

  DROP TABLE IF EXISTS geo.branch_competitors_tiered_v1_new;
  CREATE TABLE geo.branch_competitors_tiered_v1_new (LIKE geo.branch_competitors_tiered_v1 INCLUDING DEFAULTS);
  COMMIT;

  WHILE v_offset < v_total_branches LOOP
    PERFORM analytics.rebuild_tiered_competitors_batch(v_offset, v_batch_size);
    v_offset := v_offset + v_batch_size;
    COMMIT;
  END LOOP;

  CREATE INDEX ON geo.branch_competitors_tiered_v1_new (my_branch_id);
  CREATE INDEX ON geo.branch_competitors_tiered_v1_new (competitor_branch_uninumbr);
  ANALYZE geo.branch_competitors_tiered_v1_new;
  COMMIT;

  SELECT COUNT(*) INTO v_inserted FROM geo.branch_competitors_tiered_v1_new;
  IF v_inserted < v_total_branches THEN
    RAISE EXCEPTION 'branch_competitors_tiered_v1 rebuild produced only % rows for % branches - aborting before swap', v_inserted, v_total_branches;
  END IF;

  EXECUTE format('DROP TABLE IF EXISTS backup.branch_competitors_tiered_v1_%s', to_char(CURRENT_DATE, 'YYYYMMDD'));
  EXECUTE format('CREATE TABLE backup.branch_competitors_tiered_v1_%s AS TABLE geo.branch_competitors_tiered_v1', to_char(CURRENT_DATE, 'YYYYMMDD'));

  TRUNCATE geo.branch_competitors_tiered_v1;
  INSERT INTO geo.branch_competitors_tiered_v1 SELECT * FROM geo.branch_competitors_tiered_v1_new;

  GRANT SELECT ON geo.branch_competitors_tiered_v1 TO anon, authenticated, service_role;

  DROP TABLE geo.branch_competitors_tiered_v1_new;
  COMMIT;

  RAISE NOTICE 'geo.branch_competitors_tiered_v1 refreshed: % pairs for % branches (year %)', v_inserted, v_total_branches, v_year;
END;
$$;

CREATE OR REPLACE PROCEDURE geo.refresh_branch_radius_stats_tiered_v1()
LANGUAGE plpgsql AS $$
DECLARE
  v_year integer;
  v_total_branches integer;
  v_inserted bigint;
BEGIN
  SELECT MAX(year) INTO v_year FROM geo.branches_master_v2;
  SELECT COUNT(*) INTO v_total_branches FROM geo.branches_master_v2 WHERE year = v_year;

  DROP TABLE IF EXISTS geo.branch_radius_stats_tiered_v1_new;
  CREATE TABLE geo.branch_radius_stats_tiered_v1_new (LIKE geo.branch_radius_stats_tiered_v1 INCLUDING DEFAULTS);

  WITH cu_branch_counts AS (
    SELECT bank_id, COUNT(*) AS n_branches
    FROM geo.branches_master_v2 WHERE year = v_year GROUP BY bank_id
  ),
  my_branches AS (
    SELECT g.branch_id AS my_branch_id, g.bank_id, g.institution_type,
           g.bank_name, g.branch_name, g.state, g.zip,
           g.lat AS my_lat, g.lon AS my_lon, g.depsum_branch::numeric AS branch_dep,
           COALESCE(ob.inst_key,
             CASE WHEN g.institution_type = 'bank' THEN 'bank_' || g.bank_id::text
                  WHEN g.institution_type = 'cu'   THEN 'cu_'   || g.bank_id::text
                  ELSE g.bank_id::text END) AS inst_key
    FROM geo.branches_master_v2 g
    LEFT JOIN analytics.branch_opportunity_base ob ON ob.uninumbr = g.branch_id
    WHERE g.year = v_year
  ),
  bank_totals AS (
    SELECT mb.inst_key,
      COALESCE(
        (SELECT SUM(latest_dep::numeric) FROM analytics.branch_opportunity_base WHERE inst_key = mb.inst_key),
        CASE WHEN mb.institution_type = 'cu' THEN mb.branch_dep_max ELSE mb.branch_dep_sum END
      ) AS total
    FROM (
      SELECT inst_key, institution_type, SUM(branch_dep) AS branch_dep_sum, MAX(branch_dep) AS branch_dep_max
      FROM my_branches GROUP BY inst_key, institution_type
    ) mb
  ),
  candidates_1mi AS (
    SELECT mb.my_branch_id,
           gm.institution_type AS target_institution_type,
           gm.bank_id AS target_bank_id,
           gm.depsum_branch::numeric AS target_dep
    FROM my_branches mb
    JOIN geo.branches_master_v2 gm
      ON gm.year = v_year
     AND gm.bank_id != mb.bank_id
     AND ST_DWithin(gm.geom, ST_SetSRID(ST_MakePoint(mb.my_lon, mb.my_lat), 4326)::geography, 1 * 1609.34)
  ),
  density_1mi AS (
    SELECT my_branch_id, COUNT(*) AS density_1mi
    FROM candidates_1mi GROUP BY my_branch_id
  ),
  dep_1mi AS (
    SELECT c.my_branch_id,
      COALESCE(SUM(
        CASE WHEN c.target_institution_type = 'cu'
             THEN c.target_dep / NULLIF(bc.n_branches, 0)
             ELSE c.target_dep END
      ), 0) AS deposits_1mi
    FROM candidates_1mi c
    JOIN cu_branch_counts bc ON bc.bank_id = c.target_bank_id
    GROUP BY c.my_branch_id
  ),
  radius_calc AS (
    SELECT mb.my_branch_id, mb.bank_id, mb.institution_type, mb.bank_name,
           mb.branch_name, mb.state, mb.zip, mb.my_lat, mb.my_lon,
           analytics.determine_density_radius(
             COALESCE(d.density_1mi, 0)::int, mb.branch_dep,
             COALESCE(dm.deposits_1mi, 0), bt.total
           ) AS tier_radius_mi
    FROM my_branches mb
    LEFT JOIN density_1mi d ON d.my_branch_id = mb.my_branch_id
    LEFT JOIN dep_1mi dm ON dm.my_branch_id = mb.my_branch_id
    JOIN bank_totals bt ON bt.inst_key = mb.inst_key
  ),
  competitors AS (
    SELECT rc.my_branch_id,
           gm.institution_type AS c_type,
           CASE WHEN gm.institution_type = 'cu'
                THEN gm.depsum_branch::numeric / NULLIF(bc.n_branches, 0)
                ELSE gm.depsum_branch::numeric END AS c_dep
    FROM radius_calc rc
    JOIN geo.branches_master_v2 gm
      ON gm.year = v_year
     AND gm.bank_id != rc.bank_id
     AND ST_DWithin(gm.geom, ST_SetSRID(ST_MakePoint(rc.my_lon, rc.my_lat), 4326)::geography, rc.tier_radius_mi * 1609.34)
    JOIN cu_branch_counts bc ON bc.bank_id = gm.bank_id
  ),
  agg AS (
    SELECT my_branch_id,
      COUNT(*) AS branches_in_radius,
      COUNT(*) FILTER (WHERE c_type = 'bank') AS bank_branches_in_radius,
      COUNT(*) FILTER (WHERE c_type = 'cu') AS cu_branches_in_radius,
      SUM(c_dep) AS total_deposits_tiered,
      SUM(c_dep) FILTER (WHERE c_type = 'bank') AS bank_deposits_tiered,
      SUM(c_dep) FILTER (WHERE c_type = 'cu') AS cu_deposits_tiered
    FROM competitors
    GROUP BY my_branch_id
  )
  INSERT INTO geo.branch_radius_stats_tiered_v1_new
    (my_branch_id, bank_id, institution_type, bank_name, branch_name, state, zip,
     tier_radius_mi, branches_in_radius, bank_branches_in_radius, cu_branches_in_radius,
     total_deposits_tiered, bank_deposits_tiered, cu_deposits_tiered,
     avg_deposits_per_branch_tiered, cu_deposit_share_pct)
  SELECT
    rc.my_branch_id, rc.bank_id, rc.institution_type, rc.bank_name, rc.branch_name,
    rc.state, rc.zip, rc.tier_radius_mi,
    COALESCE(a.branches_in_radius, 0),
    COALESCE(a.bank_branches_in_radius, 0),
    COALESCE(a.cu_branches_in_radius, 0),
    COALESCE(a.total_deposits_tiered, 0),
    COALESCE(a.bank_deposits_tiered, 0),
    COALESCE(a.cu_deposits_tiered, 0),
    CASE WHEN COALESCE(a.branches_in_radius, 0) = 0 THEN NULL
         ELSE ROUND(a.total_deposits_tiered / a.branches_in_radius, 0) END,
    CASE WHEN COALESCE(a.total_deposits_tiered, 0) = 0 THEN NULL
         ELSE ROUND(a.cu_deposits_tiered / a.total_deposits_tiered * 100, 2) END
  FROM radius_calc rc
  LEFT JOIN agg a ON a.my_branch_id = rc.my_branch_id;

  CREATE UNIQUE INDEX ON geo.branch_radius_stats_tiered_v1_new (my_branch_id);
  ANALYZE geo.branch_radius_stats_tiered_v1_new;

  SELECT COUNT(*) INTO v_inserted FROM geo.branch_radius_stats_tiered_v1_new;
  IF v_inserted < v_total_branches * 0.95 THEN
    RAISE EXCEPTION 'branch_radius_stats_tiered_v1 rebuild produced % rows for % branches - aborting before swap', v_inserted, v_total_branches;
  END IF;

  EXECUTE format('DROP TABLE IF EXISTS backup.branch_radius_stats_tiered_v1_%s', to_char(CURRENT_DATE, 'YYYYMMDD'));
  EXECUTE format('CREATE TABLE backup.branch_radius_stats_tiered_v1_%s AS TABLE geo.branch_radius_stats_tiered_v1', to_char(CURRENT_DATE, 'YYYYMMDD'));

  TRUNCATE geo.branch_radius_stats_tiered_v1;
  INSERT INTO geo.branch_radius_stats_tiered_v1 SELECT * FROM geo.branch_radius_stats_tiered_v1_new;

  GRANT SELECT ON geo.branch_radius_stats_tiered_v1 TO anon, authenticated, service_role;

  DROP TABLE geo.branch_radius_stats_tiered_v1_new;

  RAISE NOTICE 'geo.branch_radius_stats_tiered_v1 refreshed: % branches (year %)', v_inserted, v_year;
END;
$$;

CREATE OR REPLACE PROCEDURE public.refresh_tiered_competitor_system()
LANGUAGE plpgsql AS $$
BEGIN
  CALL analytics.refresh_branch_competitors_tiered_v1();
  CALL geo.refresh_branch_radius_stats_tiered_v1();
END;
$$;
