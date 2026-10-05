"""Self-test of the swap code in ingestion/quarterly_refresh.py, run against
SCRATCH COPIES of the production tables in a throwaway schema (a84_swaptest).
Production tables are only read (to make the copies); nothing in analytics,
public or backup is written. The scratch schema is dropped at the end.

Why: the real swaps are locked (a90 hold, Coordinator sign-off), so this is how
the swap SQL, the backups, the rollback-on-failure behaviour and the guards get
exercised before anyone enables them.

  python -m analysis.a84_swap_selftest
"""

import os
import sys

import psycopg2
from psycopg2 import sql

from ingestion import quarterly_refresh as q

SCHEMA = "a84_swaptest"
results = []


def check(name, ok, detail=""):
    results.append((name, bool(ok)))
    print(f"[{'PASS' if ok else 'FAIL'}] {name}" + (f": {detail}" if detail else ""), flush=True)


def scalar(cur, query, params=None):
    cur.execute(query, params)
    return cur.fetchone()[0]


def count(cur, t):
    return scalar(cur, sql.SQL("select count(*) from {}").format(sql.Identifier(*t)))


def copy_table(cur, src, name, with_data=True):
    dst = (SCHEMA, name)
    cur.execute(sql.SQL("CREATE TABLE {} (LIKE {} INCLUDING ALL)").format(sql.Identifier(*dst), q._ident(src)))
    if with_data:
        cur.execute(sql.SQL("INSERT INTO {} SELECT * FROM {}").format(sql.Identifier(*dst), q._ident(src)))
    return dst


def fingerprint(cur, t):
    """Row count + an order-independent hash of the whole table."""
    return scalar(cur, sql.SQL("select count(*)::text || ':' || coalesce(md5(string_agg(md5(x::text), '' order by md5(x::text))), '') "
                               "from {} x").format(sql.Identifier(*t)))


def main():
    conn = psycopg2.connect(os.environ["SUPABASE_DB_URL"])
    conn.autocommit = True
    cur = conn.cursor()
    cur.execute("SET statement_timeout = '1200s'")
    cur.execute(sql.SQL("DROP SCHEMA IF EXISTS {} CASCADE").format(sql.Identifier(SCHEMA)))
    cur.execute(sql.SQL("CREATE SCHEMA {}").format(sql.Identifier(SCHEMA)))
    try:
        # ---------------- step B: snapshot swap ----------------
        T = {k: copy_table(cur, q.PROD[k], f"t_{k}") for k in ("snapshot", "latest", "ntt")}
        before = {k: fingerprint(cur, T[k]) for k in T}
        # staging = a copy of the live snapshot with every bank's cost_of_funds shifted, so the swap must visibly change data
        stg = copy_table(cur, q.PROD["snapshot"], "stg_snapshot")
        cur.execute(sql.SQL("UPDATE {} SET cost_of_funds_pct = cost_of_funds_pct + 1 WHERE cost_of_funds_pct IS NOT NULL").format(sql.Identifier(*stg)))
        stg_n = count(cur, stg)

        # validation function on scratch targets: a shifted staging must NOT reproduce live (no new quarter)
        newest = scalar(cur, sql.SQL("select max(period) from {}").format(sql.Identifier(*stg)))
        rc_banks = scalar(cur, 'select count(*) from raw."raw_schedule_RC" where period=%s', (newest,))
        c = q._validate_snapshot(cur, newest, rc_banks, targets=T, stg=stg)
        check("validation FAILS when staged cost columns differ from live with no new quarter",
              not c.ok and any("reproduce live" in n and not ok for n, ok, _ in c.rows))

        # failure injection 1: incomplete staging -> abort, nothing changed
        small = copy_table(cur, q.PROD["snapshot"], "stg_small", with_data=False)
        cur.execute(sql.SQL("INSERT INTO {} SELECT * FROM {} LIMIT 100").format(sql.Identifier(*small), sql.Identifier(*stg)))
        raised = False
        try:
            q._swap_snapshot(conn, small, T, SCHEMA, "t_small")
        except RuntimeError as e:
            raised = "staging incomplete" in str(e)
        check("incomplete staging (<40000 rows) is refused", raised)
        check("...and the scratch tables are unchanged (TRUNCATE rolled back)",
              all(fingerprint(cur, T[k]) == before[k] for k in T))

        # failure injection 2: duplicate key in staging -> INSERT fails after TRUNCATE -> whole transaction must roll back
        dup = copy_table(cur, q.PROD["snapshot"], "stg_dup")
        cur.execute(sql.SQL("ALTER TABLE {} DROP CONSTRAINT {}").format(sql.Identifier(*dup), sql.Identifier("stg_dup_pkey")))
        cur.execute(sql.SQL("INSERT INTO {} SELECT * FROM {} LIMIT 5").format(sql.Identifier(*dup), sql.Identifier(*dup)))
        raised = False
        try:
            q._swap_snapshot(conn, dup, T, SCHEMA, "t_dup")
        except psycopg2.Error:
            raised = True
        check("duplicate (inst_key, period) in staging makes the swap fail", raised)
        check("...and everything rolled back after the TRUNCATE (counts and contents identical)",
              all(fingerprint(cur, T[k]) == before[k] for k in T))
        conn.autocommit = True

        # the real swap on scratch copies
        out = q._swap_snapshot(conn, stg, T, SCHEMA, "t_ok")
        check("snapshot rows == staging rows", out["snapshot"] == stg_n, f"{out['snapshot']} vs {stg_n}")
        exp_latest = scalar(cur, sql.SQL("select count(distinct inst_key) from {}").format(sql.Identifier(*stg)))
        check("latest rows == distinct inst_keys", out["latest"] == exp_latest, f"{out['latest']} vs {exp_latest}")
        check("snapshot contents == staging contents",
              scalar(cur, sql.SQL("select count(*) from {} s join {} t using (inst_key, period) where s.cost_of_funds_pct is not distinct from t.cost_of_funds_pct").format(
                  sql.Identifier(*stg), sql.Identifier(*T["snapshot"]))) == stg_n)
        check("latest holds the newest period per bank",
              scalar(cur, sql.SQL("select count(*) from {l} l where l.period = (select max(period) from {s} where inst_key = l.inst_key)").format(
                  l=sql.Identifier(*T["latest"]), s=sql.Identifier(*stg))) == exp_latest)
        check("latest's two extra columns stay NULL",
              scalar(cur, sql.SQL("select count(marketing_expenses) + count(marketing_ratio) from {}").format(sql.Identifier(*T["latest"]))) == 0)
        check("network_top_targets copies updated from latest and rows were updated", out["ntt_updated"] > 0, f"{out['ntt_updated']} rows")
        check("...target_cof_pct now equals latest.cost_of_funds_pct for every matched row",
              scalar(cur, sql.SQL("select count(*) from {n} t join {l} l on l.inst_key = t.target_inst_key where t.target_cof_pct is distinct from l.cost_of_funds_pct").format(
                  n=sql.Identifier(*T["ntt"]), l=sql.Identifier(*T["latest"]))) == 0)
        check("...and non-financial network_top_targets columns are untouched",
              scalar(cur, sql.SQL("select count(*) from {n} t join {o} o using (my_inst_key, target_inst_key) where t.avg_vuln_score is distinct from o.avg_vuln_score or t.network_rank is distinct from o.network_rank").format(
                  n=sql.Identifier(*T["ntt"]), o=sql.Identifier(*q.PROD["ntt"]))) == 0)
        # backups hold the pre-swap data exactly
        for key, base in (("snapshot", "bank_financial_snapshot"), ("latest", "bank_financial_snapshot_latest"), ("ntt", "network_top_targets")):
            check(f"backup of {key} equals pre-swap contents", fingerprint(cur, (SCHEMA, f"{base}_t_ok")) == before[key])
        check("backup tables have RLS on", scalar(cur, "select bool_and(relrowsecurity) from pg_class where relnamespace = %s::regnamespace and relname like %s",
                                                   (SCHEMA, "%\\_t\\_ok")))
        # rollback SQL from the module docstring restores the original
        cur.execute(sql.SQL("TRUNCATE {}").format(sql.Identifier(*T["snapshot"])))
        cur.execute(sql.SQL("INSERT INTO {} SELECT * FROM {}").format(sql.Identifier(*T["snapshot"]), sql.Identifier(SCHEMA, "bank_financial_snapshot_t_ok")))
        check("documented rollback (TRUNCATE + INSERT from backup) restores the snapshot", fingerprint(cur, T["snapshot"]) == before["snapshot"])

        # ---------------- step A: UBPR layer swap ----------------
        keys = ("peer_stats", "rank", "bank_pg", "coverage")
        U = {k: copy_table(cur, q.PROD[k], f"u_{k}") for k in keys}
        ubefore = {k: fingerprint(cur, U[k]) for k in keys}
        S = {k: copy_table(cur, q.PROD[k], f"us_{k}") for k in keys}
        cur.execute(sql.SQL("DELETE FROM {} WHERE field_code = (SELECT min(field_code) FROM {})").format(sql.Identifier(*S["peer_stats"]), sql.Identifier(*S["peer_stats"])))
        cur.execute(sql.SQL("UPDATE {} SET field_value = field_value + 1 WHERE field_value IS NOT NULL").format(sql.Identifier(*S["rank"])))
        sn = {k: count(cur, S[k]) for k in keys}
        cv = q._validate_ubpr(cur, targets=U, stg=S)
        check("UBPR validation FAILS when staging is missing a field code", not cv.ok)

        # the newest bank quarter missing from staged Stats/Rank: swap-mode fails, dry-run mode only warns
        newest_raw = scalar(cur, 'select max(period) from raw."raw_UBPR"')
        M = {k: copy_table(cur, q.PROD[k], f"um_{k}") for k in keys}
        for k in ("peer_stats", "rank"):
            cur.execute(sql.SQL("DELETE FROM {} WHERE reporting_period = %s").format(sql.Identifier(*M[k])), (newest_raw,))
        strict_fail = any("includes the newest bank quarter" in n and not ok for n, ok, _ in q._validate_ubpr(cur, targets=U, stg=M, strict=True).rows)
        soft = q._validate_ubpr(cur, targets=U, stg=M, strict=False)
        soft_missing = [n for n, ok, _ in soft.rows if "newest bank quarter" in n]
        check("missing newest quarter FAILS the check in swap mode", strict_fail)
        check("...but is only a warning (no failed check about it) in dry-run mode", not soft_missing)
        check("current staged copies (live data) pass the newest-quarter check in swap mode",
              all(ok for n, ok, _ in q._validate_ubpr(cur, targets=U, stg=U, strict=True).rows if "newest bank quarter" in n))

        empty = dict(S)
        empty["coverage"] = copy_table(cur, q.PROD["coverage"], "us_empty_cov", with_data=False)
        raised = False
        try:
            q._swap_ubpr(conn, empty, U, SCHEMA, "u_empty")
        except RuntimeError as e:
            raised = "empty" in str(e)
        check("an empty staging table is refused (no truncate)", raised)
        check("...and all four scratch UBPR tables unchanged", all(fingerprint(cur, U[k]) == ubefore[k] for k in keys))

        out = q._swap_ubpr(conn, S, U, SCHEMA, "u_ok")
        check("all four UBPR tables == staging row counts", all(out[k] == sn[k] for k in keys), str(out))
        check("rank values replaced by staging values",
              scalar(cur, sql.SQL("select count(*) from {} t join {} s using (id_rssd, reporting_period, field_code, peer_group) where t.field_value is not distinct from s.field_value").format(
                  sql.Identifier(*U["rank"]), sql.Identifier(*S["rank"]))) == sn["rank"])
        for k in keys:
            check(f"backup of {k} equals pre-swap contents", fingerprint(cur, (SCHEMA, f"{U[k][1]}_u_ok")) == ubefore[k])

        # guards
        os.environ.pop(q.GUARD_ENV, None)
        for fn in (q.snapshot_swap, q.ubpr_layer_swap):
            raised = False
            try:
                fn()
            except RuntimeError as e:
                raised = "locked" in str(e)
            check(f"{fn.__name__} refuses to run without {q.GUARD_ENV}=yes", raised)

        # production untouched
        check("production snapshot/latest/ntt/ubpr tables were never written (compared with the pre-copy state)",
              all(fingerprint(cur, q.PROD[k]) == (before[k] if k in before else ubefore[k]) for k in ("snapshot", "latest", "ntt")) and
              all(fingerprint(cur, q.PROD[k]) == ubefore[k] for k in keys))
    finally:
        cur.execute(sql.SQL("DROP SCHEMA IF EXISTS {} CASCADE").format(sql.Identifier(SCHEMA)))
        conn.close()
    bad = [n for n, ok in results if not ok]
    print(f"\n{len(results) - len(bad)}/{len(results)} checks passed" + (f"; FAILED: {bad}" if bad else ""))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
