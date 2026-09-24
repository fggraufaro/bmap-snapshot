-- Mirrored copy of the year-over-year history infrastructure for git-tracked
-- readability and code review. THE DATABASE COPY IS AUTHORITATIVE -- this file is
-- kept in sync by hand whenever the live functions change via a new migration.
--
-- Applied via Supabase migration "create_bmap_year_history_tables" (2026-09-24).
--
-- Pattern: the 6 "current" tables (branches_master_v2, branch_competitors_tiered_v1,
-- branch_competitors_10mi_v2, branch_radius_stats_tiered_v1, branch_radius_stats_10mi_v2,
-- branch_opportunity_base) stay exactly as they are -- same name, same columns, same
-- refresh logic -- so no downstream app or function needs to change. Each gets a
-- companion "_history" table keyed on (natural key, year). An archive_*_history(p_year)
-- function, called once a given year's full rebuild cycle is complete, copies the
-- just-computed state into history under that year's label. Re-running archive for
-- the same year replaces only that year's slice; every other year is untouched.
--
-- 2025 was backfilled as the baseline on 2026-09-24, before any 2026 data ever
-- touched these tables, using:
--   SELECT geo.archive_branches_master_v2_history(2025);
--   SELECT geo.archive_branch_competitors_tiered_v1_history(2025);
--   SELECT geo.archive_branch_competitors_10mi_v2_history(2025);  -- clean at the time
--   SELECT geo.archive_branch_radius_stats_tiered_v1_history(2025);
--   SELECT analytics.archive_branch_opportunity_base_history(2025);
-- branch_radius_stats_10mi_v2's 2025 archive was deliberately SKIPPED at first --
-- that table had a real bug (CU branches collapsed onto one shared key) fixed later
-- the same session; back-archive it once its rebuild has actually run for a real year.
--
-- Usage once a full rebuild cycle for a year is done:
--   SELECT public.archive_bmap_year_snapshot(2026);   -- does all 6 in one call

CREATE OR REPLACE FUNCTION geo.archive_branches_master_v2_history(p_year integer)
RETURNS void LANGUAGE plpgsql AS $$
BEGIN
  DELETE FROM geo.branches_master_v2_history WHERE year = p_year;
  INSERT INTO geo.branches_master_v2_history
  SELECT *, now() FROM geo.branches_master_v2 WHERE year = p_year;
END;
$$;

CREATE OR REPLACE FUNCTION geo.archive_branch_competitors_tiered_v1_history(p_year integer)
RETURNS void LANGUAGE plpgsql AS $$
BEGIN
  DELETE FROM geo.branch_competitors_tiered_v1_history WHERE year = p_year;
  INSERT INTO geo.branch_competitors_tiered_v1_history
  SELECT *, p_year, now() FROM geo.branch_competitors_tiered_v1;
END;
$$;

-- Uses SELECT DISTINCT because the source table had exact-duplicate rows for some
-- branches at the time this was written (a separate, now-understood staleness issue,
-- not the CU key-collapse bug -- see refresh_old_10mi_competitor_system.sql).
CREATE OR REPLACE FUNCTION geo.archive_branch_competitors_10mi_v2_history(p_year integer)
RETURNS void LANGUAGE plpgsql AS $$
BEGIN
  DELETE FROM geo.branch_competitors_10mi_v2_history WHERE year = p_year;
  INSERT INTO geo.branch_competitors_10mi_v2_history
  SELECT DISTINCT *, p_year, now() FROM geo.branch_competitors_10mi_v2;
END;
$$;

CREATE OR REPLACE FUNCTION geo.archive_branch_radius_stats_tiered_v1_history(p_year integer)
RETURNS void LANGUAGE plpgsql AS $$
BEGIN
  DELETE FROM geo.branch_radius_stats_tiered_v1_history WHERE year = p_year;
  INSERT INTO geo.branch_radius_stats_tiered_v1_history
  SELECT *, p_year, now() FROM geo.branch_radius_stats_tiered_v1;
END;
$$;

CREATE OR REPLACE FUNCTION geo.archive_branch_radius_stats_10mi_v2_history(p_year integer)
RETURNS void LANGUAGE plpgsql AS $$
BEGIN
  DELETE FROM geo.branch_radius_stats_10mi_v2_history WHERE year = p_year;
  INSERT INTO geo.branch_radius_stats_10mi_v2_history
  SELECT DISTINCT *, p_year, now() FROM geo.branch_radius_stats_10mi_v2;
END;
$$;

CREATE OR REPLACE FUNCTION analytics.archive_branch_opportunity_base_history(p_year integer)
RETURNS void LANGUAGE plpgsql AS $$
BEGIN
  DELETE FROM analytics.branch_opportunity_base_history WHERE year = p_year;
  INSERT INTO analytics.branch_opportunity_base_history
  SELECT *, now() FROM analytics.branch_opportunity_base WHERE year = p_year;
END;
$$;

CREATE OR REPLACE FUNCTION public.archive_bmap_year_snapshot(p_year integer)
RETURNS void LANGUAGE plpgsql AS $$
BEGIN
  PERFORM geo.archive_branches_master_v2_history(p_year);
  PERFORM geo.archive_branch_competitors_tiered_v1_history(p_year);
  PERFORM geo.archive_branch_competitors_10mi_v2_history(p_year);
  PERFORM geo.archive_branch_radius_stats_tiered_v1_history(p_year);
  PERFORM geo.archive_branch_radius_stats_10mi_v2_history(p_year);
  PERFORM analytics.archive_branch_opportunity_base_history(p_year);
END;
$$;
