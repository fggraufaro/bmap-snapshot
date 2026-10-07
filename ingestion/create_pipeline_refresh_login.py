"""Create (or re-assert, or rotate) the restricted database login `pipeline_refresh` for the guarded refresh steps
(roadmap a33, small version). Run it yourself in a terminal that has SUPABASE_DB_URL set:

  python -m ingestion.create_pipeline_refresh_login            # create / re-assert rights, set a new password, verify
  python -m ingestion.create_pipeline_refresh_login --rotate   # new password only (rights untouched), verify
  python -m ingestion.create_pipeline_refresh_login --check    # verify only, no changes (needs the clipboard URL: see below)

Secret handling: the password is generated in memory (48 URL-safe characters), sent to Postgres as a pre-hashed
SCRAM-SHA-256 verifier (so the plaintext never reaches the server log, which records DDL), and exists afterwards only
inside the connection string that is copied to the Windows clipboard. It is never printed or written to a file.
Paste the clipboard into the Railway variable SUPABASE_REFRESH_DB_URL (mark it sealed), then clear the clipboard.

The rights come from ingestion/sql/pipeline_refresh_role.sql. Verification connects AS the new login and checks, inside
a rolled-back transaction, what it can and cannot do; nothing from the checks persists.
"""

import argparse
import base64
import hashlib
import hmac
import os
import secrets
import subprocess
import sys
from urllib.parse import quote, urlsplit, urlunsplit

import psycopg2

ROLE = "pipeline_refresh"
HERE = os.path.dirname(os.path.abspath(__file__))
SQL_FILE = os.path.join(HERE, "sql", "pipeline_refresh_role.sql")
fails = []


def check(label, ok, detail=""):
    print(f"[{'PASS' if ok else 'FAIL'}] {label}" + (f"  ({detail})" if detail and not ok else ""))
    if not ok:
        fails.append(label)


def scram_verifier(password, iterations=4096):
    salt = secrets.token_bytes(16)
    salted = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, iterations)
    client_key = hmac.new(salted, b"Client Key", hashlib.sha256).digest()
    stored_key = hashlib.sha256(client_key).digest()
    server_key = hmac.new(salted, b"Server Key", hashlib.sha256).digest()
    b64 = lambda b: base64.b64encode(b).decode()
    return f"SCRAM-SHA-256${iterations}:{b64(salt)}${b64(stored_key)}:{b64(server_key)}"


def apply_rights(admin_url):
    conn = psycopg2.connect(admin_url)
    conn.autocommit = False
    cur = conn.cursor()
    cur.execute("set lock_timeout='10s'; set statement_timeout='120s'")
    cur.execute(open(SQL_FILE, encoding="utf-8").read())
    conn.commit()
    conn.close()
    print("role and rights applied (committed)")


def set_password(admin_url, verifier):
    conn = psycopg2.connect(admin_url)
    conn.autocommit = True
    cur = conn.cursor()
    cur.execute("select count(*) from pg_roles where rolname = %s", (ROLE,))
    if not cur.fetchone()[0]:
        raise SystemExit(f"role {ROLE} does not exist; run without --rotate first")
    cur.execute(f"ALTER ROLE {ROLE} PASSWORD %s", (verifier,))
    conn.close()
    print("password set (pre-hashed)")


def verify(admin_url, new_url):
    a = psycopg2.connect(admin_url)
    a.set_session(readonly=True, autocommit=True)
    ac = a.cursor()
    ac.execute("""select rolsuper, rolcreaterole, rolcreatedb, rolbypassrls, rolcanlogin, rolconnlimit, rolreplication
                  from pg_roles where rolname = %s""", (ROLE,))
    row = ac.fetchone()
    print("pg_roles row: rolsuper=%s rolcreaterole=%s rolcreatedb=%s rolbypassrls=%s rolcanlogin=%s rolconnlimit=%s rolreplication=%s" % row)
    check("role attributes are exactly as specified", row == (False, False, False, True, True, 5, False))
    ac.execute("""select count(*) from information_schema.tables t where t.table_type = 'BASE TABLE'
                  and t.table_schema in ('raw','analytics','geo','ref','public')
                  and (has_table_privilege(%s, format('%%I.%%I', t.table_schema, t.table_name), 'INSERT')
                    or has_table_privilege(%s, format('%%I.%%I', t.table_schema, t.table_name), 'UPDATE')
                    or has_table_privilege(%s, format('%%I.%%I', t.table_schema, t.table_name), 'DELETE')
                    or has_table_privilege(%s, format('%%I.%%I', t.table_schema, t.table_name), 'TRUNCATE'))""", (ROLE,) * 4)
    check("tables the role can write = 13", ac.fetchone()[0] == 13)
    ac.execute("select count(*) from pg_class where relowner = (select oid from pg_roles where rolname = %s)", (ROLE,))
    check("the role owns nothing yet (before any step has run)", ac.fetchone()[0] == 0)
    ac.execute("""select count(*) from pg_class c join pg_namespace n on n.oid=c.relnamespace
                  where n.nspname in ('raw','analytics','geo','ref') and c.relkind in ('r','p') and pg_get_userbyid(c.relowner) <> 'postgres'
                    and c.relname not like '%%stg_pipeline%%'""")
    check("all production tables in raw/analytics/geo/ref are owned by postgres", ac.fetchone()[0] == 0)
    a.close()

    n = psycopg2.connect(new_url)
    n.autocommit = False
    c = n.cursor()
    c.execute("select current_user")
    cu = c.fetchone()[0]
    c.execute("select count(*) from analytics.bank_financial_snapshot_latest")
    rows = c.fetchone()[0]
    print(f"connected as the new login: current_user={cu}; read {rows} rows of bank_financial_snapshot_latest")
    check("connects as pipeline_refresh and reads data", cu == ROLE and rows > 0)
    c.execute("show statement_timeout"); st = c.fetchone()[0]
    check("role-level safety limits are active (statement_timeout 45min)", st in ("45min", "2700000ms", "2700s"), st)

    def attempt(label, stmt, should_work):
        c.execute("savepoint s")
        err = ""
        try:
            c.execute(stmt)
            ok = should_work
        except Exception as e:
            err = str(e).strip().splitlines()[0]
            ok = (not should_work) and ("permission denied" in err or "must be owner" in err)
        c.execute("rollback to savepoint s")
        check(label, ok, (err or "ALLOWED (BAD)") if not ok else "")

    attempt("create+drop own staging table in analytics", "create table analytics.stg_pipeline_zz_test as select 1 x; drop table analytics.stg_pipeline_zz_test", True)
    attempt("create+drop own backup table, with RLS", "create table backup.zz_test_backup as select 1 x; alter table backup.zz_test_backup enable row level security; drop table backup.zz_test_backup", True)
    attempt("write on a granted table (delete, no rows)", "delete from ref.dim_institutions where false", True)
    attempt("insert on a granted table (no rows)", "insert into raw.raw_qcew_state select * from raw.raw_qcew_state limit 0", True)
    attempt("read public.vw_bank_financial_snapshot", "select 1 from public.vw_bank_financial_snapshot limit 1", True)
    attempt("DROP a production table (bank_financial_snapshot_latest)", "drop table analytics.bank_financial_snapshot_latest", False)
    attempt("DROP raw_sod", "drop table raw.raw_sod", False)
    attempt("TRUNCATE an ungranted table (branch_opportunity_base)", "truncate analytics.branch_opportunity_base", False)
    attempt("UPDATE an ungranted table (branch_target_competitors)", "update analytics.branch_target_competitors set my_inst_key = my_inst_key where false", False)
    attempt("READ public.profiles (user emails)", "select 1 from public.profiles limit 1", False)
    attempt("READ public.resonate_audiences", "select 1 from public.resonate_audiences limit 1", False)
    attempt("CREATE a table in raw", "create table raw.zz_should_fail as select 1 x", False)
    attempt("CREATE a table in public", "create table public.zz_should_fail as select 1 x", False)
    attempt("CREATE ROLE", "create role zz_should_fail", False)
    n.rollback()
    n.close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rotate", action="store_true", help="set a new password only; leave the rights alone")
    args = ap.parse_args()
    admin_url = os.environ.get("SUPABASE_DB_URL", "")
    if not admin_url:
        raise SystemExit("SUPABASE_DB_URL is not set in this terminal")
    base = urlsplit(admin_url)
    ref = base.username.split(".", 1)[1]
    password = secrets.token_urlsafe(36)
    if not args.rotate:
        apply_rights(admin_url)
    set_password(admin_url, scram_verifier(password))
    netloc = f"{ROLE}.{ref}:{quote(password, safe='')}@{base.hostname}" + (f":{base.port}" if base.port else "")
    new_url = urlunsplit((base.scheme, netloc, base.path, base.query, ""))
    verify(admin_url, new_url)
    if fails:
        print(f"\n{len(fails)} check(s) FAILED. Nothing was copied to the clipboard. Tell Claude / the Coordinator before changing anything.")
        return 1
    subprocess.run(["clip"], input=new_url.encode("ascii"), check=True)
    print(f"\nALL CHECKS PASSED. The connection string ({len(new_url)} characters) is on your clipboard; it was not printed or saved.\n"
          "Next: paste it into Railway as SUPABASE_REFRESH_DB_URL (sealed), then clear the clipboard (Win+V -> clear all).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
