-- Mirrored copy of geo.refresh_branches_master_v2(p_year integer) for git-tracked
-- readability and code review. The DATABASE COPY IS AUTHORITATIVE -- this file is
-- kept in sync by hand whenever the live function changes via a new migration;
-- it is never deployed from here. Query pg_get_functiondef('geo.refresh_branches_master_v2'::regproc)
-- any time to confirm what's actually live.
--
-- Applied via Supabase migration "add_refresh_branches_master_v2_procedure" (2026-09-24).
--
-- Replaces the old manual staging/validate/swap script (previously run by hand from
-- a local .sql file) with one repeatable, parameterized CALL. Same field rules as
-- that script (CITYBR not CITY, RSSDID not CERT, DEPSUMBR*1000, CU composite id
-- CU_NUMBER*1e10 + JOIN_NUMBER*1e5 + SiteId, CU deposit = institution-total shares
-- repeated on every branch except "Future %" sites), but explicitly scoped to
-- p_year on both the bank (raw_sod.YEAR) and CU (raw_cu_branches/Raw_cu_fs220
-- CYCLE_DATE) sides, so it never silently mixes vintages the way an unscoped
-- "latest cycle" lookup could.
--
-- Usage:
--   CALL geo.refresh_branches_master_v2(2026);
--
-- Backs up to backup.branches_master_v2_YYYYMMDD before swapping. Aborts with no
-- changes made if either side produces zero rows, or if the total is outside
-- +/-15% of the current live row count.

CREATE OR REPLACE PROCEDURE geo.refresh_branches_master_v2(p_year integer)
LANGUAGE plpgsql AS $$
DECLARE
  v_current_count bigint;
  v_new_count bigint;
  v_bank_count bigint;
  v_cu_count bigint;
BEGIN
  SELECT COUNT(*) INTO v_current_count FROM geo.branches_master_v2;

  DROP TABLE IF EXISTS geo.branches_master_v2_new;
  CREATE TABLE geo.branches_master_v2_new (LIKE geo.branches_master_v2 INCLUDING DEFAULTS);

  -- Banks from FDIC SOD for the target year
  INSERT INTO geo.branches_master_v2_new
    (branch_id, bank_id, institution_type, year, bank_name, branch_name, address,
     city, state, zip, lat, lon, depsum_branch, insured, as_of_date, geom,
     deposit_is_estimated)
  SELECT
    s."UNINUMBR"::bigint, s."RSSDID"::bigint, 'bank', s."YEAR"::integer,
    s."NAMEFULL", s."NAMEBR", s."ADDRESBR", s."CITYBR", s."STALPBR",
    LEFT(s."ZIPBR", 5), s."SIMS_LATITUDE", s."SIMS_LONGITUDE",
    s."DEPSUMBR" * 1000, s."INSURED", make_date(s."YEAR"::integer, 6, 30),
    ST_SetSRID(ST_MakePoint(s."SIMS_LONGITUDE", s."SIMS_LATITUDE"), 4326)::geography,
    false
  FROM raw.raw_sod s
  WHERE s."YEAR" = p_year::text
    AND s."SIMS_LATITUDE" IS NOT NULL AND s."SIMS_LONGITUDE" IS NOT NULL
    AND s."STALPBR" NOT IN ('AK','HI','AS','GU','MP','PR','VI')
    AND COALESCE(s."DEPSUMBR", 0) > 0;

  -- Credit unions: latest cycle within the target year only
  WITH zip_centroids AS (
    SELECT zip, AVG(lat) AS lat, AVG(lon) AS lon
    FROM geo.branches_master_v2_new GROUP BY zip
  ),
  latest_branch_cycle AS (
    SELECT "CU_NUMBER", MAX(to_date("CYCLE_DATE",'MM/DD/YYYY HH24:MI:SS')) AS max_cycle
    FROM raw.raw_cu_branches
    WHERE EXTRACT(YEAR FROM to_date("CYCLE_DATE",'MM/DD/YYYY HH24:MI:SS')) = p_year
    GROUP BY "CU_NUMBER"
  ),
  cu_totals AS (
    SELECT f."CU_NUMBER", f."ACCT_018"::numeric AS total_shares
    FROM raw."Raw_cu_fs220" f
    JOIN latest_branch_cycle lbc
      ON lbc."CU_NUMBER" = f."CU_NUMBER"
      AND to_date(f."CYCLE_DATE",'MM/DD/YYYY HH24:MI:SS') = lbc.max_cycle
  )
  INSERT INTO geo.branches_master_v2_new
    (branch_id, bank_id, institution_type, year, bank_name, branch_name, address,
     city, state, zip, lat, lon, depsum_branch, insured, as_of_date, geom,
     deposit_is_estimated)
  SELECT
    b."CU_NUMBER"::bigint * 10000000000 + b."JOIN_NUMBER"::bigint * 100000 + b."SiteId"::bigint,
    b."CU_NUMBER"::bigint, 'cu', p_year, b."CU_NAME", b."SiteName", b."PhysicalAddressLine1",
    b."PhysicalAddressCity", b."PhysicalAddressStateCode", LEFT(b."PhysicalAddressPostalCode", 5),
    z.lat, z.lon,
    CASE WHEN b."SiteName" ILIKE 'Future %' THEN NULL ELSE t.total_shares END,
    'CU', lbc.max_cycle,
    ST_SetSRID(ST_MakePoint(z.lon, z.lat), 4326)::geography,
    CASE WHEN b."SiteName" ILIKE 'Future %' THEN false ELSE true END
  FROM raw.raw_cu_branches b
  JOIN latest_branch_cycle lbc
    ON lbc."CU_NUMBER" = b."CU_NUMBER"
    AND to_date(b."CYCLE_DATE",'MM/DD/YYYY HH24:MI:SS') = lbc.max_cycle
  JOIN zip_centroids z ON z.zip = LEFT(b."PhysicalAddressPostalCode", 5)
  LEFT JOIN cu_totals t ON t."CU_NUMBER" = b."CU_NUMBER"
  WHERE b."PhysicalAddressStateCode" NOT IN ('AK','AS','GU','HI','MP','PR','VI');

  CREATE INDEX ON geo.branches_master_v2_new USING GIST (geom);
  CREATE INDEX ON geo.branches_master_v2_new (bank_id);
  CREATE INDEX ON geo.branches_master_v2_new (institution_type);
  CREATE INDEX ON geo.branches_master_v2_new (state);
  CREATE INDEX ON geo.branches_master_v2_new (zip);
  CREATE UNIQUE INDEX ON geo.branches_master_v2_new (branch_id);
  ANALYZE geo.branches_master_v2_new;

  SELECT COUNT(*) INTO v_new_count FROM geo.branches_master_v2_new;
  SELECT COUNT(*) FILTER (WHERE institution_type='bank') INTO v_bank_count FROM geo.branches_master_v2_new;
  SELECT COUNT(*) FILTER (WHERE institution_type='cu') INTO v_cu_count FROM geo.branches_master_v2_new;

  IF v_bank_count = 0 THEN
    RAISE EXCEPTION 'branches_master_v2 rebuild for year % produced 0 bank rows - no FDIC SOD data loaded for that year? aborting before swap', p_year;
  END IF;
  IF v_cu_count = 0 THEN
    RAISE EXCEPTION 'branches_master_v2 rebuild for year % produced 0 CU rows - no NCUA cycle loaded for that year? aborting before swap', p_year;
  END IF;
  IF v_current_count > 0 AND (v_new_count < v_current_count * 0.85 OR v_new_count > v_current_count * 1.15) THEN
    RAISE EXCEPTION 'branches_master_v2 rebuild for year % produced % rows vs current % (bank=%, cu=%) - outside expected +/-15%% band, aborting before swap', p_year, v_new_count, v_current_count, v_bank_count, v_cu_count;
  END IF;

  EXECUTE format('DROP TABLE IF EXISTS backup.branches_master_v2_%s', to_char(CURRENT_DATE, 'YYYYMMDD'));
  EXECUTE format('CREATE TABLE backup.branches_master_v2_%s AS TABLE geo.branches_master_v2', to_char(CURRENT_DATE, 'YYYYMMDD'));

  TRUNCATE geo.branches_master_v2;
  INSERT INTO geo.branches_master_v2 SELECT * FROM geo.branches_master_v2_new;

  GRANT SELECT ON geo.branches_master_v2 TO anon, authenticated, service_role;

  DROP TABLE geo.branches_master_v2_new;

  RAISE NOTICE 'geo.branches_master_v2 refreshed for year %: % rows (bank=%, cu=%), was %', p_year, v_new_count, v_bank_count, v_cu_count, v_current_count;
END;
$$;
