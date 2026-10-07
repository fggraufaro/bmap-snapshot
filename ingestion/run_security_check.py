"""Run ingestion/sql/security_check.sql (roadmap a34) in a READ-ONLY session and print every section.

  python -m ingestion.run_security_check

Needs SUPABASE_DB_URL. Prints schema/table/role names and counts only, never data values.
"""

import os
import re
import sys

import psycopg2

SQL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sql", "security_check.sql")


def main():
    url = os.environ.get("SUPABASE_DB_URL", "")
    if not url:
        raise SystemExit("SUPABASE_DB_URL is not set")
    text = open(SQL_FILE, encoding="utf-8").read()
    parts = re.split(r"^-- ==== ", text, flags=re.M)[1:]
    conn = psycopg2.connect(url)
    conn.set_session(readonly=True, autocommit=True)
    cur = conn.cursor()
    cur.execute("set statement_timeout = '60s'")
    for part in parts:
        title, _, body = part.partition("\n")
        print(f"\n== {title.strip()}")
        try:
            cur.execute(body)
            cols = [d[0] for d in cur.description]
            rows = cur.fetchall()
            print(f"   {len(rows)} row(s)  columns: {', '.join(cols)}")
            for r in rows:
                print("   " + " | ".join("" if v is None else str(v) for v in r))
        except Exception as e:
            print("   ERROR:", str(e).strip().splitlines()[0])
    conn.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
