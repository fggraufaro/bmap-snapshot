-- pipeline_refresh: restricted database login for the GUARDED refresh steps (roadmap a33, small version).
-- Idempotent: safe to re-run; it re-asserts the exact set of rights (revokes, then grants). Contains no password;
-- ingestion/create_pipeline_refresh_login.py sets the password separately.
--
-- What it can do:  read the data schemas (raw, analytics, geo, ref) and two objects in public; write only the 13
--                  tables the guarded steps refresh; create its own staging/backup tables in analytics and backup.
-- What it cannot:  own or drop any production table, truncate anything outside the 13, read profiles / Resonate /
--                  persona / engagement / Rate Radar tables, create objects in raw, geo, ref or public, create roles.
-- The rebuild procedures (branches master, tiered, 10mi, opportunity base, target competitors) stay on the full login.
-- BYPASSRLS is needed because every table has RLS on with no policy; it is harmless here because the grants are narrow.

DO $$
BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'pipeline_refresh') THEN
    CREATE ROLE pipeline_refresh LOGIN NOSUPERUSER NOCREATEDB NOCREATEROLE NOREPLICATION BYPASSRLS CONNECTION LIMIT 5;
  END IF;
END $$;

-- (the superuser / bypassrls / replication attributes cannot be re-asserted by a non-superuser; they are set at creation
--  and checked by create_pipeline_refresh_login.py)
ALTER ROLE pipeline_refresh LOGIN CONNECTION LIMIT 5;
ALTER ROLE pipeline_refresh SET statement_timeout = '45min';                    -- steps set their own; this is the ceiling
ALTER ROLE pipeline_refresh SET lock_timeout = '2min';
ALTER ROLE pipeline_refresh SET idle_in_transaction_session_timeout = '30min';

-- start from nothing, then grant exactly what is needed
REVOKE ALL ON ALL TABLES IN SCHEMA raw, analytics, geo, ref, public, backup FROM pipeline_refresh;
REVOKE ALL ON ALL SEQUENCES IN SCHEMA raw, analytics, geo, ref, public, backup FROM pipeline_refresh;
REVOKE ALL ON ALL FUNCTIONS IN SCHEMA raw, analytics, geo, ref, public, backup FROM pipeline_refresh;
REVOKE ALL ON SCHEMA raw, analytics, geo, ref, public, backup FROM pipeline_refresh;

GRANT USAGE ON SCHEMA raw, analytics, geo, ref, public, backup TO pipeline_refresh;
GRANT CREATE ON SCHEMA analytics, backup TO pipeline_refresh;       -- own staging and backup tables only

GRANT SELECT ON ALL TABLES IN SCHEMA raw, analytics, geo, ref TO pipeline_refresh;
-- views run with their owner's rights, so SELECT on the view is enough
GRANT SELECT ON public.network_top_targets, public.vw_bank_financial_snapshot TO pipeline_refresh;

GRANT INSERT, UPDATE, DELETE, TRUNCATE ON
  analytics.bank_financial_snapshot, analytics.bank_financial_snapshot_latest,
  analytics.ubpr_peer_stats_clean, analytics.ubpr_rank_clean,
  analytics.ubpr_bank_peer_group, analytics.ubpr_rank_coverage,
  ref.dim_institutions, public.network_top_targets,
  raw.raw_cbp_totals, raw.raw_business_formation_state, raw.raw_irs_migration_state,
  raw.raw_qcew_state, raw.raw_cfpb_complaints_trend
TO pipeline_refresh;

GRANT EXECUTE ON FUNCTION public.populate_network_top_targets() TO pipeline_refresh;
