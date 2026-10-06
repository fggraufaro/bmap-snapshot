"""Control check: is every table and view in the database covered by a refresh plan?

The plan lives in ingestion/table_coverage.json (one row per object: status, how it is refreshed,
owner, cadence, what to do if it is not covered). This script compares that file with the live
database and fails if they disagree, so a new table cannot appear without someone deciding how it
gets refreshed.

  python -m ingestion.table_coverage            # run the checks, print the report
  python -m ingestion.table_coverage --write    # also write ingestion/TABLE_COVERAGE.md (+ .json snapshot for the page)

Read-only: opens a read-only session and never writes to the database.

Status values
  AUTO     refreshed by a command-center button (or by a procedure that button calls)
  LOCKED   a button exists but the destructive part is gated (ALLOW_ANALYTICS_SWAP / sign-off)
  HELD     a button exists and is deliberately not to be clicked
  ARCHIVE  history table written by the yearly archive button
  SERVICE  refreshed by another service (Rate Radar, persona layer, Resonate)
  APP      written by an application (job queue, users, engagement tool); not pipeline data
  DERIVED  a view; follows its base tables
  STATIC   reference data loaded once, no refresh path (needs an owner/decision if a tool reads it)
  GAP      feeds a tool or a rebuild and has no refresh path
  UNUSED   nothing reads it and nothing refreshes it (keep or drop)
  SYSTEM   PostGIS / platform objects
The backup schema is exempt from the manifest but checked for row-level security.
"""

import argparse
import json
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import date

import psycopg2

HERE = os.path.dirname(os.path.abspath(__file__))
MANIFEST = os.path.join(HERE, "table_coverage.json")
SCHEMAS = ("public", "analytics", "geo", "raw", "ref", "pbi")
NEEDS_ACTION = {"GAP", "STATIC", "UNUSED"}
NOT_REFRESHED = {"GAP", "STATIC", "UNUSED", "SERVICE", "HELD"}
# Procedures whose inputs must all be covered. Rebuild chain, in run order.
PROCS = ["geo.refresh_branches_master_v2", "analytics.rebuild_tiered_radius_batch",
         "analytics.refresh_branch_competitors_tiered_v1", "geo.refresh_branch_competitors_10mi_v2",
         "analytics.refresh_branch_opportunity_base", "analytics.refresh_branch_target_competitors",
         "public.populate_network_top_targets"]
# Newest-data date per table, shown in the report (live values, never stored in the manifest).
ASOF = {
    "raw.raw_sod": 'select max("YEAR")::text from raw.raw_sod',
    "raw.raw_cu_branches": "select max(to_date(\"CYCLE_DATE\",'MM/DD/YYYY HH24:MI:SS'))::text from raw.raw_cu_branches",
    "raw.Raw_cu_fs220": "select max(to_date(\"CYCLE_DATE\",'MM/DD/YYYY HH24:MI:SS'))::text from raw.\"Raw_cu_fs220\"",
    "raw.raw_income": 'select max("YEAR")::text from raw.raw_income',
    "raw.raw_population": 'select max("YEAR")::text from raw.raw_population',
    "raw.raw_occupation": 'select max("YEAR")::text from raw.raw_occupation',
    "raw.raw_zhvi": "select max(period)::text from raw.raw_zhvi",
    "raw.raw_schedule_RI": "select max(period)::text from raw.\"raw_schedule_RI\"",
    "raw.raw_schedule_RC": "select max(period)::text from raw.\"raw_schedule_RC\"",
    "raw.raw_UBPR": "select max(period)::text from raw.\"raw_UBPR\"",
    "raw.raw_cfpb_complaints": "select max(loaded_at)::date::text from raw.raw_cfpb_complaints",
    "raw.raw_cfpb_complaints_trend": "select max(loaded_at)::date::text from raw.raw_cfpb_complaints_trend",
    "raw.raw_rate_radar": "select max(run_date)::text from raw.raw_rate_radar",
    "public.rate_observations": "select max(observed_at)::date::text from public.rate_observations",
    "public.persona_runs": "select max(run_date)::date::text from public.persona_runs",
    "public.resonate_audiences": "select max(fetched_at)::date::text from public.resonate_audiences",
    "geo.branches_master_v2": "select max(as_of_date)::text from geo.branches_master_v2",
    "geo.branch_competitors_10mi_v2": "select max(as_of_date)::text from geo.branch_competitors_10mi_v2",
    "analytics.branch_opportunity_base": "select max(year)::text from analytics.branch_opportunity_base",
    "analytics.branch_opportunity_base_history": "select string_agg(distinct year::text, ', ' order by year::text) from analytics.branch_opportunity_base_history",
    "analytics.bank_financial_snapshot_latest": "select max(period)::text from analytics.bank_financial_snapshot_latest",
    "analytics.bank_financial_snapshot": "select max(period)::text from analytics.bank_financial_snapshot",
}


def _session():
    url = os.environ.get("SUPABASE_DB_URL", "")
    if not url:
        raise SystemExit("SUPABASE_DB_URL is not set")
    conn = psycopg2.connect(url)
    conn.set_session(readonly=True, autocommit=True)
    cur = conn.cursor()
    cur.execute("set statement_timeout = '120s'")
    return conn, cur


def _live(cur):
    cur.execute("""select n.nspname, c.relname, c.relkind::text, c.relrowsecurity, c.oid
                   from pg_class c join pg_namespace n on n.oid = c.relnamespace
                   where c.relkind in ('r','p','v','m') and n.nspname = any(%s)""", (list(SCHEMAS),))
    live = {f"{s}.{n}": dict(kind=k, rls=rls, oid=o) for s, n, k, rls, o in cur.fetchall()}
    cur.execute("""select c.relname, c.relrowsecurity from pg_class c join pg_namespace n on n.oid = c.relnamespace
                   where c.relkind in ('r','p') and n.nspname = 'backup'""")
    backups = cur.fetchall()
    cur.execute("""select distinct table_schema || '.' || table_name from information_schema.role_table_grants
                   where grantee in ('anon','authenticated','PUBLIC') and privilege_type = 'SELECT'
                   and table_schema = any(%s)""", (list(SCHEMAS),))
    readable = {r[0] for r in cur.fetchall()}
    for k in live:
        live[k]["anon_read"] = k in readable
    return live, backups


def _view_bases(cur, live):
    """view -> set of base tables, resolved through nested views."""
    cur.execute("""select vc.oid, d.refobjid from pg_depend d
                   join pg_rewrite r on r.oid = d.objid join pg_class vc on vc.oid = r.ev_class and vc.relkind in ('v','m')
                   join pg_class rc on rc.oid = d.refobjid and rc.relkind in ('r','p','v','m')
                   where d.refobjid <> vc.oid""")
    deps = defaultdict(set)
    for v, r in cur.fetchall():
        deps[v].add(r)
    name = {m["oid"]: k for k, m in live.items()}

    def bases(oid, seen):
        if oid in seen or oid not in name:
            return set()
        seen.add(oid)
        if live[name[oid]]["kind"] in ("r", "p"):
            return {name[oid]}
        out = set()
        for d in deps.get(oid, ()):
            out |= bases(d, seen)
        return out
    return {k: bases(m["oid"], set()) for k, m in live.items() if m["kind"] in ("v", "m")}


def _proc_inputs(cur, live):
    cur.execute("""select n.nspname || '.' || p.proname, lower(p.prosrc) from pg_proc p
                   join pg_namespace n on n.oid = p.pronamespace where n.nspname = any(%s)""", (list(SCHEMAS),))
    src = dict(cur.fetchall())
    out = {}
    for proc in PROCS:
        text = src.get(proc)
        if text is None:
            out[proc] = None
            continue
        out[proc] = sorted(k for k in live if re.search(
            r'(?<![a-z0-9_.])(?:' + k.split(".")[0] + r'\.)?"?' + re.escape(k.split(".")[1].lower()) + r'"?(?![a-z0-9_])', text))
    return out


def _counts(cur, live):
    rows = {}
    for k, m in live.items():
        if m["kind"] not in ("r", "p"):
            continue
        s, n = k.split(".", 1)
        try:
            cur.execute(f'select count(*) from "{s}"."{n}"')
            rows[k] = cur.fetchone()[0]
        except Exception:
            rows[k] = None
    return rows


def _asof(cur):
    out = {}
    for k, q in ASOF.items():
        try:
            cur.execute(q)
            out[k] = cur.fetchone()[0]
        except Exception as e:
            out[k] = "n/a"
    return out


def run(write=False):
    from ingestion.pipeline_steps import STEP_BY_ID

    manifest = json.load(open(MANIFEST, encoding="utf-8"))
    by_obj = {r["object"]: r for r in manifest}
    conn, cur = _session()
    live, backups = _live(cur)
    bases = _view_bases(cur, live)
    procs = _proc_inputs(cur, live)
    counts = _counts(cur, live)
    asof = _asof(cur)
    conn.close()

    errors, warnings = [], []
    if len(by_obj) != len(manifest):
        errors.append("duplicate objects in the manifest: " + ", ".join(k for k, c in Counter(r["object"] for r in manifest).items() if c > 1))
    for k in sorted(set(live) - set(by_obj)):
        errors.append(f"NOT IN MANIFEST: {k} ({'view' if live[k]['kind'] in 'vm' else 'table'}) exists in the database with no refresh decision")
    for k in sorted(set(by_obj) - set(live)):
        errors.append(f"STALE MANIFEST ROW: {k} is in the manifest but not in the database")
    for r in manifest:
        for sid in re.findall(r"button: (\w+)", r["mechanism"]):
            if sid not in STEP_BY_ID:
                errors.append(f"{r['object']}: mechanism names step '{sid}', which does not exist in pipeline_steps")
        if r["status"] in NEEDS_ACTION and not r["action"]:
            errors.append(f"{r['object']}: status {r['status']} needs an action (what is the decision/fix?)")
        if r["status"] == "DERIVED" and r["object"] in live and live[r["object"]]["kind"] not in "vm":
            errors.append(f"{r['object']}: marked DERIVED but is a table")
        if r["status"] not in ("DERIVED", "SYSTEM") and r["object"] in live and live[r["object"]]["kind"] in "vm":
            errors.append(f"{r['object']}: is a view but not marked DERIVED")
    for k, m in live.items():
        if m["kind"] in ("r", "p") and not m["rls"] and by_obj.get(k, {}).get("status") != "SYSTEM":
            if m["anon_read"]:
                warnings.append(f"RLS OFF AND READABLE BY anon/authenticated: {k} (exposed now)")
            else:
                warnings.append(f"RLS off, no anon/authenticated grant: {k} (not exposed today; enable RLS so a later GRANT cannot expose it)")
    for n, rls in backups:
        if not rls:
            warnings.append(f"RLS OFF: backup.{n}")
    # a tool-feeding table must have a refresh path
    for r in manifest:
        if r["tools"] and r["status"] in NOT_REFRESHED - {"SERVICE"}:
            warnings.append(f"FEEDS A TOOL, NO PIPELINE REFRESH: {r['object']} [{r['status']}] -> {', '.join(r['tools'])}")
    # views inherit staleness
    view_risk = {}
    for v, bs in bases.items():
        risky = sorted(b for b in bs if by_obj.get(b, {}).get("status") in {"GAP", "STATIC"})
        view_risk[v] = risky
        if risky and by_obj.get(v, {}).get("status") == "DERIVED":
            warnings.append(f"VIEW DEPENDS ON UNREFRESHED TABLE: {v} <- {', '.join(risky)}")
    # rebuild-chain inputs
    proc_rows = {}
    for proc, ins in procs.items():
        if ins is None:
            errors.append(f"PROCEDURE NOT FOUND: {proc} (update PROCS in table_coverage.py)")
            continue
        proc_rows[proc] = [(i, by_obj.get(i, {}).get("status", "MISSING")) for i in ins]
        for i, st in proc_rows[proc]:
            if st in {"GAP", "STATIC", "UNUSED", "MISSING"}:
                warnings.append(f"REBUILD INPUT NOT REFRESHED: {proc} reads {i} [{st}]")

    status_counts = Counter(r["status"] for r in manifest)
    print(f"Objects in database: {len(live)} ({sum(1 for m in live.values() if m['kind'] in 'rp')} tables, "
          f"{sum(1 for m in live.values() if m['kind'] in 'vm')} views); in manifest: {len(by_obj)}; backup-schema tables: {len(backups)}")
    print("Status counts:", dict(sorted(status_counts.items())))
    print(f"\nERRORS ({len(errors)}):")
    for e in errors:
        print("  " + e)
    print(f"\nWARNINGS ({len(warnings)}):")
    for w in warnings:
        print("  " + w)

    if write:
        _write_md(manifest, live, counts, asof, bases, proc_rows, errors, warnings, status_counts, len(backups))
    return 1 if errors else 0


def _write_md(manifest, live, counts, asof, bases, proc_rows, errors, warnings, status_counts, n_backup):
    order = ["GAP", "STATIC", "UNUSED", "LOCKED", "HELD", "AUTO", "ARCHIVE", "SERVICE", "APP", "DERIVED", "SYSTEM"]
    L = [f"# Table coverage control document", "",
         f"Generated {date.today().isoformat()} by `python -m ingestion.table_coverage --write` from the live database "
         f"and `ingestion/table_coverage.json`. Do not edit by hand: edit the JSON, re-run, commit both.", "",
         f"**Result: {'FAIL' if errors else 'PASS'}** - {len(errors)} errors, {len(warnings)} warnings. "
         f"{len(live)} objects in the database ({len(manifest) - status_counts['DERIVED']} tables, {status_counts['DERIVED']} views) "
         f"all have a row below; the backup schema ({n_backup} tables) is exempt.", "",
         "## Summary", "", "| Status | Objects | Meaning |", "|---|---|---|"]
    meaning = {"AUTO": "Refreshed by a command-center button", "LOCKED": "Button exists, destructive part gated",
               "HELD": "Button exists, deliberately not clicked", "ARCHIVE": "History, yearly archive button",
               "SERVICE": "Refreshed by another service", "APP": "Written by an application",
               "DERIVED": "View, follows its base tables", "STATIC": "Loaded once, no refresh path",
               "GAP": "Feeds a tool or rebuild, no refresh path", "UNUSED": "No reader, no refresh: keep or drop",
               "SYSTEM": "PostGIS / platform"}
    for s in order:
        L.append(f"| {s} | {status_counts.get(s, 0)} | {meaning[s]} |")
    L += ["", "## Errors", ""] + ([f"- {e}" for e in errors] or ["None."])
    L += ["", "## Warnings (open items)", ""] + ([f"- {w}" for w in warnings] or ["None."])
    L += ["", "## Rebuild chain: inputs of each procedure", "", "| Procedure | Reads (status) |", "|---|---|"]
    for p, ins in proc_rows.items():
        L.append(f"| `{p}` | " + ", ".join(f"`{i}` {s}" for i, s in ins if not i.endswith("_history")) + " |")
    L += ["", "## Every object", ""]
    for s in order:
        rows = [r for r in manifest if r["status"] == s]
        if not rows:
            continue
        L += [f"### {s} ({len(rows)})", "", "| Object | Rows | Newest data | How it is refreshed | Owner | Cadence | Feeds | Action / note |", "|---|---|---|---|---|---|---|---|"]
        for r in sorted(rows, key=lambda x: x["object"]):
            k = r["object"]
            n = counts.get(k)
            rows_txt = f"{n:,}" if isinstance(n, int) else "view"
            extra = r["action"] or ""
            if r["note"]:
                extra = (extra + " - " if extra else "") + r["note"]
            if r["status"] == "DERIVED" and bases.get(k):
                extra = (extra + " " if extra else "") + "Base: " + ", ".join(sorted(bases[k]))
            L.append(f"| `{k}` | {rows_txt} | {asof.get(k, '')} | {r['mechanism']} | {r['owner']} | {r['cadence']} | "
                     f"{', '.join(r['tools'])} | {extra} |")
        L.append("")
    open(os.path.join(HERE, "TABLE_COVERAGE.md"), "w", encoding="utf-8").write("\n".join(L) + "\n")
    snap = dict(generated=date.today().isoformat(), errors=errors, warnings=warnings, status_counts=status_counts,
                backup_tables=n_backup, procs={p: ins for p, ins in proc_rows.items()},
                rows=[dict(r, rows=counts.get(r["object"]), asof=asof.get(r["object"]), base=sorted(bases.get(r["object"], []))) for r in manifest])
    json.dump(snap, open(os.path.join(HERE, "table_coverage_snapshot.json"), "w", encoding="utf-8"), indent=1)
    print("\nwrote ingestion/TABLE_COVERAGE.md and table_coverage_snapshot.json")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--write", action="store_true")
    sys.exit(run(ap.parse_args().write))
