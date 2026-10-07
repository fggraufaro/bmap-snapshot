-- Recurring Supabase security check (roadmap a34). READ-ONLY: only SELECTs on catalogs. No values, no secrets.
-- Run:  python -m ingestion.run_security_check      (opens a read-only session; prints each section)
-- Sections are separated by lines starting with "-- ==== ". "Expected" says what a clean result looks like.

-- ==== 1. Tables with row-level security OFF (non-system schemas). Expected: only PostGIS public.spatial_ref_sys (a92) and nothing else.
SELECT n.nspname AS schema, c.relname AS "table", c.relkind
FROM pg_class c JOIN pg_namespace n ON n.oid = c.relnamespace
WHERE c.relkind IN ('r', 'p') AND NOT c.relrowsecurity
  AND n.nspname NOT IN ('pg_catalog', 'information_schema') AND n.nspname NOT LIKE 'pg\_toast%' AND n.nspname NOT LIKE 'pg\_temp%'
ORDER BY 1, 2;

-- ==== 2. anon / authenticated / PUBLIC holding any privilege beyond SELECT on a table or view. Expected: none.
SELECT n.nspname AS schema, c.relname AS object, c.relkind,
       CASE a.grantee WHEN 0 THEN 'PUBLIC' ELSE a.grantee::regrole::text END AS grantee,
       string_agg(a.privilege_type, ', ' ORDER BY a.privilege_type) AS privileges
FROM pg_class c JOIN pg_namespace n ON n.oid = c.relnamespace
CROSS JOIN LATERAL aclexplode(COALESCE(c.relacl, acldefault('r', c.relowner))) a
WHERE c.relkind IN ('r', 'p', 'v', 'm', 'f') AND a.privilege_type <> 'SELECT'
  AND (a.grantee = 0 OR a.grantee::regrole::text IN ('anon', 'authenticated'))
  AND n.nspname NOT IN ('pg_catalog', 'information_schema') AND n.nspname NOT LIKE 'pg\_toast%'
GROUP BY 1, 2, 3, 4
ORDER BY 1, 2;

-- ==== 3. Policies that apply to anon (directly or through PUBLIC). Expected: none, unless a table deliberately serves anon.
SELECT schemaname AS schema, tablename AS "table", policyname, cmd, permissive, roles
FROM pg_policies
WHERE roles && ARRAY['anon', 'public']::name[]
ORDER BY 1, 2, 3;

-- ==== 3b. All policies in the app schemas, for reference (count per table).
SELECT schemaname AS schema, tablename AS "table", count(*) AS policies
FROM pg_policies GROUP BY 1, 2 ORDER BY 1, 2;

-- ==== 4a. Tables created in the last 30 days (named in recent migrations) that exist now, with RLS state and policy count. Expected: rls_on true for all.
WITH created AS (
  SELECT DISTINCT lower(replace(x[1], '"', '')) AS name
  FROM supabase_migrations.schema_migrations m,
       regexp_matches(array_to_string(m.statements, ' '), 'create\s+table\s+(?:if\s+not\s+exists\s+)?([a-z_0-9."]+)', 'gi') x
  WHERE to_timestamp(substr(m.version, 1, 14), 'YYYYMMDDHH24MISS') > now() - interval '30 days')
SELECT n.nspname AS schema, c.relname AS "table", c.relrowsecurity AS rls_on,
       (SELECT count(*) FROM pg_policies p WHERE p.schemaname = n.nspname AND p.tablename = c.relname) AS policies
FROM created cr
JOIN pg_namespace n ON true
JOIN pg_class c ON c.relnamespace = n.oid AND c.relkind IN ('r', 'p')
 AND (lower(n.nspname || '.' || c.relname) = cr.name OR (position('.' IN cr.name) = 0 AND lower(c.relname) = cr.name AND n.nspname = 'public'))
ORDER BY c.relrowsecurity, 1, 2;

-- ==== 4b. Tables in the backup schema whose name carries a date suffix from the last 30 days, with RLS state.
SELECT n.nspname AS schema, c.relname AS "table", c.relrowsecurity AS rls_on
FROM pg_class c JOIN pg_namespace n ON n.oid = c.relnamespace
WHERE c.relkind IN ('r', 'p') AND n.nspname = 'backup'
  AND substring(c.relname FROM '(20[0-9]{6})') IS NOT NULL
  AND to_date(substring(c.relname FROM '(20[0-9]{6})'), 'YYYYMMDD') > current_date - 30
ORDER BY 2;

-- ==== 4c. RLS state of every table whose name suggests a recent, dated copy anywhere (backup/bak/pre_/_20 in the name) outside the backup schema.
SELECT n.nspname AS schema, c.relname AS "table", c.relrowsecurity AS rls_on
FROM pg_class c JOIN pg_namespace n ON n.oid = c.relnamespace
WHERE c.relkind IN ('r', 'p') AND n.nspname NOT IN ('backup', 'pg_catalog', 'information_schema')
  AND (c.relname ~ 'backup|_bak|_pre_|_20[0-9]{6}')
ORDER BY 1, 2;

-- ==== 5. SECURITY DEFINER functions executable by anon or authenticated. Expected: only the Hub radius lookups (public.branches_within_radius, _batch).
SELECT n.nspname AS schema, p.proname AS function, pg_get_userbyid(p.proowner) AS owner,
       has_function_privilege('anon', p.oid, 'EXECUTE') AS anon_can_run,
       has_function_privilege('authenticated', p.oid, 'EXECUTE') AS authenticated_can_run
FROM pg_proc p JOIN pg_namespace n ON n.oid = p.pronamespace
WHERE p.prosecdef AND p.prokind IN ('f', 'p')
  AND n.nspname NOT IN ('pg_catalog', 'information_schema')
  AND (has_function_privilege('anon', p.oid, 'EXECUTE') OR has_function_privilege('authenticated', p.oid, 'EXECUTE'))
ORDER BY 1, 2;

-- ==== 6. Views that bypass row-level security: owner is a superuser or has BYPASSRLS, no security_invoker, and anon/authenticated can SELECT.
SELECT n.nspname AS schema, c.relname AS view, pg_get_userbyid(c.relowner) AS owner,
       has_table_privilege('anon', c.oid, 'SELECT') AS anon_can_select,
       has_table_privilege('authenticated', c.oid, 'SELECT') AS authenticated_can_select
FROM pg_class c JOIN pg_namespace n ON n.oid = c.relnamespace
JOIN pg_roles r ON r.oid = c.relowner
WHERE c.relkind IN ('v', 'm')
  AND NOT COALESCE('security_invoker=true' = ANY (c.reloptions) OR 'security_invoker=on' = ANY (c.reloptions), false)
  AND (r.rolsuper OR r.rolbypassrls)
  AND (has_table_privilege('anon', c.oid, 'SELECT') OR has_table_privilege('authenticated', c.oid, 'SELECT'))
  AND n.nspname NOT IN ('pg_catalog', 'information_schema', 'extensions', 'graphql', 'graphql_public', 'vault', 'pgsodium', 'pgsodium_masks')
ORDER BY 1, 2;

-- ==== 7. pipeline_refresh login: attributes, objects owned, tables with write rights. Expected: 0 owned, 13 write tables, no super/createrole/createdb/replication, bypassrls true.
SELECT r.rolname, r.rolsuper, r.rolcreaterole, r.rolcreatedb, r.rolreplication, r.rolbypassrls, r.rolcanlogin, r.rolconnlimit,
       (SELECT count(*) FROM pg_class c WHERE c.relowner = r.oid) AS objects_owned,
       (SELECT count(*) FROM information_schema.tables t
         WHERE t.table_type = 'BASE TABLE' AND t.table_schema IN ('raw', 'analytics', 'geo', 'ref', 'public', 'backup')   -- app schemas; cron.job_run_details is DELETE-able by every role via PUBLIC (Supabase default)
           AND (has_table_privilege(r.rolname, format('%I.%I', t.table_schema, t.table_name), 'INSERT')
             OR has_table_privilege(r.rolname, format('%I.%I', t.table_schema, t.table_name), 'UPDATE')
             OR has_table_privilege(r.rolname, format('%I.%I', t.table_schema, t.table_name), 'DELETE')
             OR has_table_privilege(r.rolname, format('%I.%I', t.table_schema, t.table_name), 'TRUNCATE'))) AS tables_with_write_rights
FROM pg_roles r WHERE r.rolname = 'pipeline_refresh';
