"""Refresh ref.dim_institutions (the institution directory) from the newest raw
SOD (banks) and NCUA branch file (credit unions).

Why: the table was bulk-loaded once (around 2026-04-02) and has no refresh path,
so every institution that appeared, or was renamed, since then is missing or
stale -- found 2026-10-06: 73 banks and 68 credit unions in the current raw data
had no row, plus about 95 banks whose names changed (e.g. "The Bank of Orrick" is
now "TBO Bank"). Readers: the Hub / Growth Map / Opportunity View institution
search, Rate Radar's typeahead, public.populate_network_top_targets
(my_institution), vw_prospecting_score, vw_bank_directory and others.

How each column is derived (checked against the existing rows before writing
this; the display rule holds for all 8,846 rows):
  banks  (raw.raw_sod, newest YEAR, one row per RSSDID, highest-deposit branch row):
         inst_key 'bank_'||RSSDID, institution_id = RSSDID, institution_name =
         NAMEFULL, state_hq = STALP, city_hq = CITY (HQ city), insured_code = INSURED
  cus    (raw.raw_cu_branches, newest CYCLE_DATE, main-office row):
         inst_key 'cu_'||CU_NUMBER, institution_id = CU_NUMBER, institution_name =
         CU_NAME, state_hq/city_hq = main office physical state/city, insured_code 'CU'
  institution_name_display = institution_name || ' -- ' || state_hq

What it changes (and what it never does):
  - ADDS institutions that are in the newest raw data but not in the table.
  - RENAMES an existing institution only when the new name differs in words (after
    ignoring case and punctuation). Case-only differences (about 430 banks) are left
    alone, and state/city/insured of existing rows are never touched.
  - Institutions headquartered in AK, AS, GU, HI, MP, PR or VI are NOT added unless they
    have at least one branch in geo.branches_master_v2: the whole pipeline excludes those
    states by BRANCH location, so an HQ-in-Hawaii bank with no scored branches would show
    up in the Hub search with no data behind it, while a Hawaii-HQ union with branches in
    scored states (cu_10882, cu_7989) must be named. (The table already holds 10 such
    rows from the original load; they are left as they are.) First dry run: 70 of the 141
    missing institutions were in those states. Because this reads the branch master, the
    step runs right after the branch-master rebuild.
  - NEVER deletes (institutions that disappear, e.g. after a merger, stay, because
    archives and old target rows still reference their keys).

DRY RUN vs APPLY: applying needs ALLOW_ANALYTICS_SWAP=yes on the service (same lock as
the other production swaps; not set until sign-off). Without it the step does a dry
run: it computes the plan, validates it, prints every rename and a sample of adds, and
changes nothing. That keeps the step safe inside "Run all". Apply = back up the table
to backup.dim_institutions_<ts> (RLS on), then INSERT + UPDATE in ONE transaction with
counts and unchanged-row checks inside it; any mismatch rolls everything back.

ROLLBACK (after an apply): one transaction,
    TRUNCATE ref.dim_institutions;
    INSERT INTO ref.dim_institutions SELECT * FROM backup.dim_institutions_<ts>;

CLI:
  python -m ingestion.refresh_dim_institutions            # dry run unless the unlock is set
  python -m ingestion.refresh_dim_institutions --dry-run
"""

import datetime
import os
import sys

import psycopg2
from psycopg2 import sql

from ingestion.quarterly_refresh import GUARD_ENV, Checks, _backup, _ident, _log, _scalar, _session

PROD = ("ref", "dim_institutions")
NORM = "lower(regexp_replace({}, '[^a-zA-Z0-9]', '', 'g'))"

DESIRED = """
select * from (
  with sod_year as (select max("YEAR") y from raw.raw_sod),
  banks as (
    select distinct on (s."RSSDID")
           'bank_' || s."RSSDID" as inst_key, s."RSSDID"::bigint as institution_id,
           btrim(s."NAMEFULL") as institution_name, upper(btrim(s."STALP")) as state_hq,
           btrim(s."CITY") as city_hq, btrim(s."INSURED") as insured_code, 'bank'::text as institution_type
    from raw.raw_sod s, sod_year
    where s."YEAR" = sod_year.y and s."RSSDID" ~ '^[0-9]+$' and nullif(btrim(s."NAMEFULL"), '') is not null
    order by s."RSSDID", s."DEPSUMBR"::numeric desc nulls last
  ),
  cu_cycle as (select max(to_date("CYCLE_DATE", 'MM/DD/YYYY HH24:MI:SS')) c from raw.raw_cu_branches),
  cus as (
    select distinct on (r."CU_NUMBER")
           'cu_' || r."CU_NUMBER" as inst_key, r."CU_NUMBER"::bigint as institution_id,
           btrim(r."CU_NAME") as institution_name, upper(btrim(r."PhysicalAddressStateCode")) as state_hq,
           btrim(r."PhysicalAddressCity") as city_hq, 'CU'::text as insured_code, 'cu'::text as institution_type
    from raw.raw_cu_branches r, cu_cycle
    where to_date(r."CYCLE_DATE", 'MM/DD/YYYY HH24:MI:SS') = cu_cycle.c
      and r."CU_NUMBER" ~ '^[0-9]+$' and nullif(btrim(r."CU_NAME"), '') is not null
    order by r."CU_NUMBER", (r."MainOffice" = 'Yes') desc
  )
  select *, institution_name || ' -- ' || state_hq as institution_name_display
  from (select * from banks union all select * from cus) x
  where nullif(state_hq, '') is not null
    and (state_hq not in ('AK', 'AS', 'GU', 'HI', 'MP', 'PR', 'VI')
         or exists (select 1 from geo.branches_master_v2 g
                    where g.institution_type = x.institution_type and g.bank_id = x.institution_id))
) desired
"""


def _plan_sql(target):
    t = _ident(target)
    n_new, n_old = NORM.format("d.institution_name"), NORM.format("l.institution_name")
    return sql.SQL(
        "select d.*, l.institution_name as old_name, "
        "(l.inst_key is null) as is_add, "
        "(l.inst_key is not null and " + n_new + " <> " + n_old + ") as is_rename "
        "from (" + DESIRED + ") d left join {t} l on l.inst_key = d.inst_key"
    ).format(t=t)


def _plan(cur, target):
    cur.execute(_plan_sql(target))
    cols = [c.name for c in cur.description]
    rows = [dict(zip(cols, r)) for r in cur.fetchall()]
    return ([r for r in rows if r["is_add"]], [r for r in rows if r["is_rename"]], len(rows))


def _validate(cur, adds, renames, desired_n, target):
    c = Checks()
    live_n = _scalar(cur, sql.SQL("select count(*) from {}").format(_ident(target)))
    live_by_type = dict(cur.execute(sql.SQL("select institution_type, count(*) from {} group by 1").format(_ident(target))) or cur.fetchall())
    cur.execute(sql.SQL("select count(*) filter (where institution_name_display = institution_name || ' -- ' || state_hq), count(*) from {}").format(_ident(target)))
    ok_disp, tot = cur.fetchone()
    c.add("existing rows follow the display rule (name -- STATE)", ok_disp == tot, f"{ok_disp}/{tot}")
    cur.execute("select count(*) filter (where inst_key like 'bank\\_%'), count(*) filter (where inst_key like 'cu\\_%') from (" + DESIRED + ") d")
    d_banks, d_cus = cur.fetchone()
    c.add("newest SOD covers >= 90% of the banks in the table", d_banks >= 0.9 * live_by_type.get("bank", 0), f"{d_banks} vs {live_by_type.get('bank', 0)}")
    c.add("newest NCUA file covers >= 90% of the credit unions in the table", d_cus >= 0.9 * live_by_type.get("cu", 0), f"{d_cus} vs {live_by_type.get('cu', 0)}")
    c.add("adds <= 5% of the table", len(adds) <= 0.05 * live_n, f"{len(adds)} of {live_n}")
    c.add("renames <= 5% of the table", len(renames) <= 0.05 * live_n, f"{len(renames)} of {live_n}")
    bad = [r["inst_key"] for r in adds if not (r["institution_name"] and r["state_hq"] and r["institution_id"] and r["institution_type"])]
    c.add("every added row has name, state, id and type", not bad, f"{len(bad)} incomplete")
    dup = _scalar(cur, sql.SQL("select count(*) - count(distinct inst_key) from (" + DESIRED + ") d"))
    c.add("no duplicate inst_key in the desired set", dup == 0, f"{dup} duplicates")
    return c, live_n


def _report(adds, renames, desired_n):
    by = lambda rows: {t: sum(1 for r in rows if r["institution_type"] == t) for t in ("bank", "cu")}
    _log(f"  desired institutions: {desired_n} | to add: {len(adds)} {by(adds)} | to rename: {len(renames)} {by(renames)}")
    _log("  renames (existing name -> new name):")
    for r in sorted(renames, key=lambda r: r["inst_key"]):
        _log(f"    {r['inst_key']}: {r['old_name']!r} -> {r['institution_name']!r}")
    _log("  sample of institutions to add:")
    for r in sorted(adds, key=lambda r: r["inst_key"])[:12]:
        _log(f"    {r['inst_key']}: {r['institution_name']} ({r['city_hq']}, {r['state_hq']}, {r['insured_code']})")


def dry_run(target=PROD):
    _log("Institution directory refresh: DRY RUN (no table is changed).")
    with _session("refresh_dim_institutions") as conn, conn.cursor() as cur:
        cur.execute("SET SESSION CHARACTERISTICS AS TRANSACTION READ ONLY")
        adds, renames, desired_n = _plan(cur, target)
        _report(adds, renames, desired_n)
        checks, _ = _validate(cur, adds, renames, desired_n, target)
        checks.raise_if_failed()
        _log("DRY RUN PASSED. Enable the swap (ALLOW_ANALYTICS_SWAP=yes) to apply it.")


def apply(target=PROD, backup_schema="backup", suffix=None):
    suffix = suffix or datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    _log(f"Institution directory refresh: APPLY (backup suffix {suffix}).")
    with _session("refresh_dim_institutions") as conn:
        with conn.cursor() as cur:
            adds, renames, desired_n = _plan(cur, target)
            _report(adds, renames, desired_n)
            checks, live_n = _validate(cur, adds, renames, desired_n, target)
            checks.raise_if_failed()
            backup = _backup(cur, target, backup_schema, target[1], suffix)
        rename_keys = [r["inst_key"] for r in renames]
        conn.autocommit = False
        try:
            with conn.cursor() as cur:
                cur.execute("SET LOCAL lock_timeout = '30s'")
                t = _ident(target)
                n_new, n_old = NORM.format("d.institution_name"), NORM.format("t.institution_name")
                cur.execute(sql.SQL(
                    "INSERT INTO {t} (inst_key, institution_id, institution_name, institution_name_display, state_hq, city_hq, insured_code, institution_type) "
                    "select d.inst_key, d.institution_id, d.institution_name, d.institution_name_display, d.state_hq, d.city_hq, d.insured_code, d.institution_type "
                    "from (" + DESIRED + ") d where not exists (select 1 from {t} x where x.inst_key = d.inst_key)").format(t=t))
                inserted = cur.rowcount
                cur.execute(sql.SQL(
                    "UPDATE {t} t set institution_name = d.institution_name, institution_name_display = d.institution_name || ' -- ' || t.state_hq "
                    "from (" + DESIRED + ") d where t.inst_key = d.inst_key and " + n_new + " <> " + n_old).format(t=t))
                updated = cur.rowcount
                if inserted != len(adds):
                    raise RuntimeError(f"inserted {inserted} rows but the plan had {len(adds)} adds")
                if updated != len(renames):
                    raise RuntimeError(f"updated {updated} rows but the plan had {len(renames)} renames")
                new_n = _scalar(cur, sql.SQL("select count(*) from {}").format(t))
                if new_n != live_n + inserted:
                    raise RuntimeError(f"row count {new_n}, expected {live_n + inserted}")
                changed_other = _scalar(cur, sql.SQL("select count(*) from {t} t join {b} b using (inst_key) where t::text <> b::text and not (t.inst_key = any(%s))")
                                        .format(t=t, b=_ident(backup)), (rename_keys,))
                if changed_other:
                    raise RuntimeError(f"{changed_other} rows outside the plan changed")
                bad_disp = _scalar(cur, sql.SQL("select count(*) from {} where institution_name_display <> institution_name || ' -- ' || state_hq").format(t))
                if bad_disp:
                    raise RuntimeError(f"{bad_disp} rows break the display rule")
            conn.commit()
        except BaseException:
            conn.rollback()
            raise
        finally:
            conn.autocommit = True
    _log(f"  APPLIED: inserted {inserted}, renamed {updated}, table now {new_n} rows. Backup: {backup[0]}.{backup[1]}")
    return {"inserted": inserted, "renamed": updated, "rows": new_n}


def refresh_dim_institutions():
    """Command-center entry point: applies when unlocked, otherwise a safe dry run."""
    if os.environ.get(GUARD_ENV) == "yes":
        apply()
    else:
        _log(f"{GUARD_ENV} is not set: running as a DRY RUN only.")
        dry_run()


if __name__ == "__main__":
    dry_run() if "--dry-run" in sys.argv else refresh_dim_institutions()
