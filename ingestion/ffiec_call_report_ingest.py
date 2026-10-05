"""Loads one FFIEC bank quarter into raw."raw_schedule_RI", raw."raw_schedule_RC"
and raw."raw_UBPR" (a85).

Source: the same public bulk portal as the UBPR peer-stats loads (see
ffiec_bulk_portal.py for the postback mechanics) -- a14's 2026-09-11
conclusion that this portal is "unautomatable" was wrong: it's a plain
ASP.NET form that scripts fine (proven on UBPR Stats/Rank, 2026-10-02).
  - RI / RC:  product "Call Reports -- Single Period"
              (ReportingSeriesSinglePeriod), files
              "FFIEC CDR Call Schedule RI/RC MMDDYYYY.txt"
  - raw_UBPR: "UBPR Ratio -- Four Periods" (the Single Period variant
              returns an HTML error page for 06/30/2026), file
              "...UBPR Ratios Summary Ratios YYYY.txt" -- the existing 45
              raw_UBPR columns are exactly that file's columns.

Format matches existing rows: every value text, NULL for empty cells, RIAD/
RCON codes as published (thousands of dollars), plus the manual-era "DATE"
text column (written MM/DD/YYYY) and the real DATE "period" column.
raw_UBPR also gets "Reporting Period" in FFIEC's own string form
("6/30/2026 11:59:59 PM") -- earlier loads left period/Reporting Period NULL
and needed fix_ubpr_q1_2026_period / patch_ubpr_null_periods; this script
always sets both.

SAFETY: insert-only. The target tables have no unique key, so a re-run
would duplicate -- the script refuses to touch a table that already has any
row for the period (it never deletes). Never calls refresh_bmap_after_upload.

Usage:
  python -m ingestion.ffiec_call_report_ingest --latest --dry-run     # what the command-center step does
  python -m ingestion.ffiec_call_report_ingest --period 06/30/2026 --dry-run
  python -m ingestion.ffiec_call_report_ingest --period 06/30/2026
  python -m ingestion.ffiec_call_report_ingest --period 06/30/2025 --only ri rc   # raw_UBPR 6/30/2025 already loaded
"""

import argparse
import csv
import io
import sys
from datetime import datetime

import requests

from ingestion.ffiec_bulk_portal import _scrape_form_state, fetch_zip
from ingestion.supabase_client import count, get, insert

CALL_PRODUCT = "ReportingSeriesSinglePeriod"
UBPR_PRODUCT = "PerformanceReportingSeriesFourPeriods"


def _rows(raw_bytes):
    text = raw_bytes.decode("utf-8", errors="replace")
    return list(csv.reader(io.StringIO(text), delimiter="\t"))


def _table_columns(table):
    sample = get(table, schema="raw", params="select=*&limit=1")
    if not sample:
        raise RuntimeError(f"raw.{table} is empty -- can't infer its columns")
    return set(sample[0].keys())


def _loaded_rows(table, period_iso):
    return count(table, schema="raw", params=f"period=eq.{period_iso}")


def latest_quarter():
    """Newest quarter-end the portal offers for the Call Reports product
    (dropdown is listed newest first), as MM/DD/YYYY."""
    _, _, options = _scrape_form_state(requests.Session(), CALL_PRODUCT)
    return options[0][1]


def parse_schedule(raw_bytes, table_cols, period_iso, date_text):
    """Call Report schedule file: row 1 = codes ("IDRSSD" is quoted),
    row 2 = descriptions, data from row 3. Trailing empty header column
    ignored."""
    rows = _rows(raw_bytes)
    codes = [c.strip().strip('"') for c in rows[0]]
    keep = [(i, c) for i, c in enumerate(codes) if c and c in table_cols]
    dropped = [c for c in codes if c and c not in table_cols]
    out = []
    for rec in rows[2:]:
        if not rec or not rec[0].strip():
            continue
        row = {c: (rec[i].strip() or None) if i < len(rec) else None for i, c in keep}
        row["DATE"] = date_text
        row["period"] = period_iso
        out.append(row)
    return out, dropped


def parse_ubpr(raw_bytes, table_cols, period_dt, period_iso):
    """UBPR Ratios file: row 1 = codes, rows 2-3 = long name/description,
    data from row 4. Keeps only the target quarter's rows (the Four Periods
    zip carries every quarter of the year)."""
    rows = _rows(raw_bytes)
    codes = [c.strip() for c in rows[0]]
    keep = [(i, c) for i, c in enumerate(codes) if c and c in table_cols]
    dropped = [c for c in codes if c and c not in table_cols]
    out = []
    for rec in rows[3:]:
        if not rec or not rec[0].strip():
            continue
        try:
            got = datetime.strptime(rec[0].strip(), "%m/%d/%Y %I:%M:%S %p").date()
        except ValueError:
            continue
        if got != period_dt:
            continue
        row = {c: (rec[i].strip() or None) if i < len(rec) else None for i, c in keep}
        row["period"] = period_iso
        out.append(row)
    return out, dropped


def _report(name, rows, id_col, dropped):
    banks = len({r[id_col] for r in rows})
    print(f"  {name}: {len(rows)} rows, {banks} distinct banks"
          f"{'' if not dropped else f'  (file columns not in table, ignored: {len(dropped)})'}")
    if len(rows) != banks:
        raise RuntimeError(f"{name}: {len(rows)} rows but {banks} distinct banks -- unexpected duplicates in source file")


def load(period_str, only=None, dry_run=False, skip_loaded=False):
    """skip_loaded=False (CLI default): refuse if any target already has the
    period. skip_loaded=True (command-center step): a table that already has
    the period is skipped IF its row count equals the source file's -- a
    mismatch (e.g. a crashed partial insert) raises instead of silently
    skipping or duplicating. Never deletes either way."""
    period_dt = datetime.strptime(period_str, "%m/%d/%Y").date()
    period_iso = period_dt.isoformat()
    date_text = period_dt.strftime("%m/%d/%Y")
    targets = set(only) if only else {"ri", "rc", "ubpr"}

    names = {"ri": "raw_schedule_RI", "rc": "raw_schedule_RC", "ubpr": "raw_UBPR"}
    existing = {k: _loaded_rows(names[k], period_iso) for k in targets}
    if not skip_loaded:
        for key in sorted(targets):
            if existing[key]:
                raise RuntimeError(f"raw.{names[key]} already has {existing[key]} rows for {period_iso} -- refusing to "
                                   f"insert (no unique key; would duplicate). Nothing was written.")

    parsed = {}  # key -> (table, rows)
    if targets & {"ri", "rc"}:
        zf, _ = fetch_zip(CALL_PRODUCT, period_str)
        mmddyyyy = period_dt.strftime("%m%d%Y")
        for key, sched in (("ri", "RI"), ("rc", "RC")):
            if key not in targets:
                continue
            raw = zf.read(f"FFIEC CDR Call Schedule {sched} {mmddyyyy}.txt")
            rows, dropped = parse_schedule(raw, _table_columns(names[key]), period_iso, date_text)
            _report(names[key], rows, "IDRSSD", dropped)
            parsed[key] = (names[key], rows)

    if "ubpr" in targets:
        zf, _ = fetch_zip(UBPR_PRODUCT, period_dt.year)
        raw = zf.read(f"FFIEC CDR UBPR Ratios Summary Ratios {period_dt.year}.txt")
        rows, dropped = parse_ubpr(raw, _table_columns("raw_UBPR"), period_dt, period_iso)
        for r in rows:
            r["Reporting Period"] = f"{period_dt.month}/{period_dt.day}/{period_dt.year} 11:59:59 PM"
        if rows:
            _report("raw_UBPR", rows, "ID RSSD", dropped)
        parsed["ubpr"] = ("raw_UBPR", rows)

    plan = []
    for key in sorted(parsed):
        table, rows = parsed[key]
        if not rows:
            msg = f"raw.{table}: {period_iso} not in the source file yet (UBPR can lag Call Reports)"
            if skip_loaded:
                print(f"  {msg} -- skipping.")
                continue
            raise RuntimeError(msg)
        if existing[key]:
            if existing[key] != len(rows):
                raise RuntimeError(f"raw.{table} has {existing[key]} rows for {period_iso} but the source has "
                                   f"{len(rows)} -- possible partial load. Refusing to insert or skip; needs a manual look.")
            print(f"  raw.{table}: {period_iso} already loaded ({existing[key]} rows, matches source) -- skipping.")
            continue
        plan.append((table, rows))

    if dry_run:
        print(f"Dry run -- nothing written ({len(plan)} table(s) would load).")
        return
    for table, rows in plan:
        sent = insert(table, rows, schema="raw", batch_size=500)
        print(f"Inserted {sent} rows into raw.{table} ({period_iso}).")
    if not plan:
        print(f"Nothing to load for {period_iso}.")


def load_latest():
    """Command-center entry point: load the newest quarter the portal has.
    Raw tables only -- does NOT run refresh_bmap_after_upload (a database
    procedure that truncates/rebuilds the analytics snapshot and score tables;
    nothing in this pipeline calls it)."""
    period = latest_quarter()
    print(f"Latest FFIEC quarter on the portal: {period}")
    load(period, skip_loaded=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--period", default=None, help="quarter-end as MM/DD/YYYY, e.g. 06/30/2026 "
                     "(omit with --latest)")
    ap.add_argument("--latest", action="store_true", help="load the newest quarter the portal offers, skipping tables that already have it")
    ap.add_argument("--only", nargs="+", choices=["ri", "rc", "ubpr"], default=None,
                     help="subset of tables (default: all three)")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()
    if args.latest:
        period = latest_quarter()
        print(f"Latest FFIEC quarter on the portal: {period}")
        load(period, args.only, args.dry_run, skip_loaded=True)
    elif args.period:
        load(args.period, args.only, args.dry_run)
    else:
        ap.error("give --period MM/DD/YYYY or --latest")


if __name__ == "__main__":
    sys.exit(main())
