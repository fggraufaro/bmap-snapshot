"""Self-test of ingestion/refresh_dim_institutions.py against a SCRATCH COPY of
ref.dim_institutions in a throwaway schema (dimtest). Production is only read;
the scratch schema is dropped at the end.

  python -m analysis.dim_refresh_selftest
"""

import os
import sys

import psycopg2
from psycopg2 import sql

from ingestion import refresh_dim_institutions as r
from ingestion.quarterly_refresh import GUARD_ENV

SCHEMA = "dimtest"
results = []


def check(name, ok, detail=""):
    results.append((name, bool(ok)))
    print(f"[{'PASS' if ok else 'FAIL'}] {name}" + (f": {detail}" if detail else ""), flush=True)


def scalar(cur, query, params=None):
    cur.execute(query, params)
    return cur.fetchone()[0]


def fp(cur, t):
    return scalar(cur, sql.SQL("select count(*)::text || ':' || coalesce(md5(string_agg(md5(x::text), '' order by md5(x::text))), '') from {} x").format(sql.Identifier(*t)))


def main():
    conn = psycopg2.connect(os.environ["SUPABASE_DB_URL"])
    conn.autocommit = True
    cur = conn.cursor()
    cur.execute(sql.SQL("DROP SCHEMA IF EXISTS {} CASCADE").format(sql.Identifier(SCHEMA)))
    cur.execute(sql.SQL("CREATE SCHEMA {}").format(sql.Identifier(SCHEMA)))
    T = (SCHEMA, "dim")
    prod_before = fp(cur, r.PROD)
    try:
        cur.execute(sql.SQL("CREATE TABLE {} (LIKE ref.dim_institutions INCLUDING ALL)").format(sql.Identifier(*T)))
        cur.execute(sql.SQL("INSERT INTO {} SELECT * FROM ref.dim_institutions").format(sql.Identifier(*T)))
        before = fp(cur, T)
        adds, renames, desired_n = r._plan(cur, T)
        check("plan finds adds and renames on the real data", len(adds) > 0 and len(renames) > 0, f"{len(adds)} adds, {len(renames)} renames")
        check("no add is an existing key", all(a["old_name"] is None for a in adds))
        check("case-only name differences are NOT renames",
              all(r_["old_name"].lower() != r_["institution_name"].lower() for r_ in renames))
        c, live_n = r._validate(cur, adds, renames, desired_n, T)
        check("validation passes on the real data", c.ok)

        # dry run changes nothing
        r.dry_run(T)
        check("dry run leaves the scratch table identical", fp(cur, T) == before)

        # guard: without the unlock, the command-center entry point only dry-runs
        os.environ.pop(GUARD_ENV, None)
        r.refresh_dim_institutions()
        check("refresh_dim_institutions() without the unlock changes nothing (production untouched)", fp(cur, r.PROD) == prod_before)

        # failure injection: a constraint that rejects one of the added rows -> everything rolls back
        victim = sorted(adds, key=lambda a: a["inst_key"])[len(adds) // 2]["institution_name"]
        cur.execute(sql.SQL("ALTER TABLE {} ADD CONSTRAINT no_victim CHECK (institution_name <> %s)").format(sql.Identifier(*T)), (victim,))
        raised = False
        try:
            r.apply(T, SCHEMA, "t_fail")
        except psycopg2.Error:
            raised = True
        check("a row the table rejects makes the apply fail", raised)
        check("...and the table is unchanged (rolled back), backup was still taken", fp(cur, T) == before and scalar(cur, sql.SQL("select count(*) from {}").format(sql.Identifier(SCHEMA, "dim_t_fail"))) == live_n)
        cur.execute(sql.SQL("ALTER TABLE {} DROP CONSTRAINT no_victim").format(sql.Identifier(*T)))

        # the real apply on the scratch copy
        out = r.apply(T, SCHEMA, "t_ok")
        check("inserted == planned adds", out["inserted"] == len(adds), str(out))
        check("renamed == planned renames", out["renamed"] == len(renames))
        check("row count == before + adds", out["rows"] == live_n + len(adds))
        check("every desired institution now exists", scalar(cur, sql.SQL("select count(*) from (" + r.DESIRED + ") d where not exists (select 1 from {} x where x.inst_key = d.inst_key)").format(sql.Identifier(*T))) == 0)
        check("renamed rows carry the new name and a consistent display name",
              scalar(cur, sql.SQL("select count(*) from {} t join (" + r.DESIRED + ") d using (inst_key) where t.inst_key = any(%s) and (t.institution_name <> d.institution_name or t.institution_name_display <> d.institution_name || ' -- ' || t.state_hq)").format(sql.Identifier(*T)),
                     ([x["inst_key"] for x in renames],)) == 0)
        check("display rule holds for every row", scalar(cur, sql.SQL("select count(*) from {} where institution_name_display <> institution_name || ' -- ' || state_hq").format(sql.Identifier(*T))) == 0)
        check("rows outside the plan are byte-identical to before",
              scalar(cur, sql.SQL("select count(*) from {t} t join {b} b using (inst_key) where t::text <> b::text and not (t.inst_key = any(%s))").format(t=sql.Identifier(*T), b=sql.Identifier(SCHEMA, "dim_t_ok")),
                     ([x["inst_key"] for x in renames],)) == 0)
        check("no row was deleted", scalar(cur, sql.SQL("select count(*) from {b} b where not exists (select 1 from {t} t where t.inst_key = b.inst_key)").format(t=sql.Identifier(*T), b=sql.Identifier(SCHEMA, "dim_t_ok"))) == 0)
        check("backup equals the pre-apply table", fp(cur, (SCHEMA, "dim_t_ok")) == before)
        check("backup table has RLS on", scalar(cur, "select relrowsecurity from pg_class where oid = %s::regclass", (f"{SCHEMA}.dim_t_ok",)))

        # idempotent: a second run has nothing left to do
        a2, r2, _ = r._plan(cur, T)
        check("a second run finds nothing to do", len(a2) == 0 and len(r2) == 0, f"{len(a2)} adds, {len(r2)} renames")
        after = fp(cur, T)
        out2 = r.apply(T, SCHEMA, "t_again")
        check("...and applying again changes nothing", out2["inserted"] == 0 and out2["renamed"] == 0 and fp(cur, T) == after)

        # documented rollback restores the pre-apply table
        cur.execute(sql.SQL("TRUNCATE {}").format(sql.Identifier(*T)))
        cur.execute(sql.SQL("INSERT INTO {} SELECT * FROM {}").format(sql.Identifier(*T), sql.Identifier(SCHEMA, "dim_t_ok")))
        check("documented rollback restores the table", fp(cur, T) == before)

        # the apply entry point with the unlock set operates on PRODUCTION, so it is NOT exercised here.
        check("production ref.dim_institutions was never written", fp(cur, r.PROD) == prod_before)
    finally:
        cur.execute(sql.SQL("DROP SCHEMA IF EXISTS {} CASCADE").format(sql.Identifier(SCHEMA)))
        conn.close()
    bad = [n for n, ok in results if not ok]
    print(f"\n{len(results) - len(bad)}/{len(results)} checks passed" + (f"; FAILED: {bad}" if bad else ""))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
