"""Self-test for ingestion/session2_datasets.py. Runs the APPLY paths against throwaway scratch schemas
(copies of the live tables); production tables are only read, and a before/after fingerprint proves they
were not written. The CFPB loader is tested with a mocked API (the real one refuses this machine).

  python -m analysis.session2_datasets_selftest
"""

import hashlib
import os
import sys
from datetime import date

import psycopg2

from ingestion import session2_datasets as s

SCHEMA, BAK = "s2test", "s2test_bak"
TABLES = ["raw_cbp_totals", "raw_business_formation_state", "raw_irs_migration_state", "raw_qcew_state",
          "raw_cfpb_complaints", "raw_cfpb_complaints_trend"]
fails = 0


def check(label, ok, detail=""):
    global fails
    print(f"[{'PASS' if ok else 'FAIL'}] {label}" + (f"  ({detail})" if detail and not ok else ""))
    fails += (not ok)


def fingerprint(cur, schema, table):
    cur.execute(f'select count(*), md5(string_agg(t::text, \'|\' order by t::text)) from {schema}."{table}" t')
    return cur.fetchone()


def main():
    conn = psycopg2.connect(os.environ["SUPABASE_DB_URL"])
    conn.autocommit = True
    cur = conn.cursor()
    prod = {t: fingerprint(cur, "raw", t) for t in TABLES}
    try:
        cur.execute(f"drop schema if exists {SCHEMA} cascade; drop schema if exists {BAK} cascade")
        cur.execute(f"create schema {SCHEMA}; create schema {BAK}")
        for t in TABLES:
            cur.execute(f'create table {SCHEMA}."{t}" (like raw."{t}" including all)')
            cur.execute(f'insert into {SCHEMA}."{t}" select * from raw."{t}"')

        # 1. guard: with the env var unset, refresh() is a dry run and changes nothing
        os.environ.pop(s.GUARD_ENV, None)
        before = fingerprint(cur, SCHEMA, "raw_irs_migration_state")
        s.refresh("irs_migration", schema=SCHEMA, backup_schema=BAK)
        check("no ALLOW_DATASET_LOAD: dry run leaves the table unchanged", fingerprint(cur, SCHEMA, "raw_irs_migration_state") == before)

        # 2. apply each replace-mode loader against the scratch copy
        for name in ("cbp", "business_formation", "irs_migration", "qcew"):
            spec = s.SPECS[name]
            pre = fingerprint(cur, SCHEMA, spec["table"])
            diff = s.refresh(name, schema=SCHEMA, backup_schema=BAK, apply=True)
            rows = s._live_rows(cur, SCHEMA, spec)
            new, _ = spec["fetch"](cur, SCHEMA)
            d2 = s._compare(spec, rows, new)
            check(f"{name}: after apply the table equals the source ({d2['new']} rows)", d2["changed"] == d2["added"] == d2["removed"] == 0 and d2["live"] == d2["new"])
            cur.execute("select table_name from information_schema.tables where table_schema=%s and table_name like %s", (BAK, spec["table"] + "_%"))
            baks = [r[0] for r in cur.fetchall()]
            check(f"{name}: a backup table was created", len(baks) == 1)
            if baks:
                check(f"{name}: the backup equals the pre-apply table", fingerprint(cur, BAK, baks[0])[0] == pre[0])
                cur.execute("select relrowsecurity from pg_class c join pg_namespace n on n.oid=c.relnamespace where n.nspname=%s and c.relname=%s", (BAK, baks[0]))
                check(f"{name}: the backup has RLS on", cur.fetchone()[0] is True)

        cur.execute(f"select tax_year_pair, count(*), min(people_in) > 0 from {SCHEMA}.raw_irs_migration_state group by 1")
        r = cur.fetchall()
        check("irs: exactly one vintage, 51 states", len(r) == 1 and r[0][1] == 51, str(r))

        # 3. failure paths leave the table untouched
        spec = s.SPECS["business_formation"]
        base = fingerprint(cur, SCHEMA, spec["table"])
        good, _ = spec["fetch"](cur, SCHEMA)

        def refuse(label, rows, msg_part):
            try:
                s.refresh("business_formation", schema=SCHEMA, backup_schema=BAK, apply=True, fetch=lambda c, sc: (rows, "test"))
                check(label, False, "no error raised")
            except Exception as e:
                check(label, msg_part in str(e) and fingerprint(cur, SCHEMA, spec["table"]) == base, str(e)[:120])

        refuse("duplicate keys are refused", good + [dict(good[0])], "duplicate keys")
        refuse("a table more than 10% smaller is refused", good[:40], "smaller")
        refuse("no rows is refused", [], "no rows")
        bad = [dict(r_) for r_ in good]
        bad[10]["ba2024"] = "not a number"
        refuse("a bad value mid-insert rolls the whole reload back", bad, "")

        # 4. CFPB trend: upsert, with a mocked API
        spec = s.SPECS["cfpb_trend"]
        calls = []

        def fake_count(name, start, end):
            calls.append((name, start, end))
            return (len(name) + (end - start).days) % 17

        def fetch(cur_, schema_):
            return s._fetch_cfpb_trend(cur_, schema_, today=date(2026, 10, 6), count=fake_count, calibrate=False)

        pre = fingerprint(cur, SCHEMA, spec["table"])[0]
        s.refresh("cfpb_trend", schema=SCHEMA, backup_schema=BAK, apply=True, fetch=fetch)
        cur.execute(f"select week_ending, count(*) from {SCHEMA}.raw_cfpb_complaints_trend group by 1 order by 1")
        weeks = cur.fetchall()
        check("cfpb: new weeks (2026-09-24, 2026-10-01) were added and the stored weeks kept",
              [str(w) for w, _ in weeks] == ["2026-09-10", "2026-09-17", "2026-09-24", "2026-10-01"] and all(n == 53 for _, n in weeks), str(weeks))
        check("cfpb: the window is the 90 days ending the Thursday", all((e - st).days == 89 and e.weekday() == 3 for _, st, e in calls))
        s.refresh("cfpb_trend", schema=SCHEMA, backup_schema=BAK, apply=True, fetch=fetch)
        check("cfpb: re-running the same week is idempotent", fingerprint(cur, SCHEMA, spec["table"])[0] == pre + 106)
        cur.execute(f"""select count(*) from {SCHEMA}.raw_cfpb_complaints_trend t join {SCHEMA}.raw_cfpb_complaints c using (inst_key)
                        where t.week_ending >= '2026-09-24' and t.insufficient_volume <> (c.complaints_2024 < 10)""")
        check("cfpb: insufficient_volume follows 'fewer than 10 complaints in 2024'", cur.fetchone()[0] == 0)
        cur.execute(f"select count(*) from {SCHEMA}.raw_cfpb_complaints_trend t join raw.raw_cfpb_complaints_trend p using (inst_key, week_ending) "
                    "where t.trailing_90d_count <> p.trailing_90d_count or t.insufficient_volume <> p.insufficient_volume")
        check("cfpb: the stored weeks 2026-09-10 and 2026-09-17 are untouched", cur.fetchone()[0] == 0)
        # calibration: a mock API that matches the 90-day window must be reported as the better fit
        cur.execute(f"select inst_key, trailing_90d_count from {SCHEMA}.raw_cfpb_complaints_trend where week_ending = (select max(week_ending) from {SCHEMA}.raw_cfpb_complaints_trend)")
        stored = dict(cur.fetchall())
        cur.execute(f"select inst_key, cfpb_company_name, complaints_2024 from {SCHEMA}.raw_cfpb_complaints order by inst_key")
        insts = cur.fetchall()
        byname = {n: k for k, n, _ in insts}
        import contextlib, io
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            s._cfpb_calibrate(cur, SCHEMA, insts, lambda name, a, b: stored[byname[name]] + (0 if (b - a).days == 89 else 1))
        out = buf.getvalue()
        check("cfpb: calibration reports the matching window (53 of 53) and the wrong one (0 of 53)",
              "start = week - 89): 53 of 53" in out and "start = week - 90): 0 of 53" in out, out)
        try:
            import requests
            real = requests.get
            requests.get = lambda *a, **k: type("R", (), {"status_code": 403})()
            try:
                s._cfpb_count("X", date(2026, 1, 1), date(2026, 3, 1))
                check("cfpb: a 403 raises a clear error", False)
            except RuntimeError as e:
                check("cfpb: a 403 raises a clear error", "403" in str(e))
            finally:
                requests.get = real
        except Exception as e:
            check("cfpb: a 403 raises a clear error", False, str(e))
    finally:
        cur.execute(f"drop schema if exists {SCHEMA} cascade; drop schema if exists {BAK} cascade")
        after = {t: fingerprint(cur, "raw", t) for t in TABLES}
        check("production raw tables were never written", after == prod)
        conn.close()
    print(f"\n{'ALL PASSED' if not fails else str(fails) + ' FAILED'}")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
