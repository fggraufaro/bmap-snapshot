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
  python -m ingestion.ffiec_call_report_ingest --period 06/30/2026 --dry-run
  python -m ingestion.ffiec_call_report_ingest --period 06/30/2026
  python -m ingestion.ffiec_call_report_ingest --period 06/30/2025 --only ri rc   # raw_UBPR 6/30/2025 already loaded
"""

import argparse
import csv
import io
import sys
from datetime import datetime

from ingestion.ffiec_bulk_portal import fetch_zip
from ingestion.supabase_client import get, insert

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


def _already_loaded(table, period_iso):
    rows = get(table, schema="raw", params=f"select=period&period=eq.{period_iso}&limit=1")
    return bool(rows)


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


def load(period_str, only=None, dry_run=False):
    period_dt = datetime.strptime(period_str, "%m/%d/%Y").date()
    period_iso = period_dt.isoformat()
    date_text = period_dt.strftime("%m/%d/%Y")
    targets = set(only) if only else {"ri", "rc", "ubpr"}

    # Fail fast, before any download, if a target already has the period.
    names = {"ri": "raw_schedule_RI", "rc": "raw_schedule_RC", "ubpr": "raw_UBPR"}
    for key in sorted(targets):
        if _already_loaded(names[key], period_iso):
            raise RuntimeError(f"raw.{names[key]} already has rows for {period_iso} -- refusing to insert "
                               f"(no unique key; would duplicate). Nothing was written.")

    plan = []  # (table, rows)
    if targets & {"ri", "rc"}:
        zf, _ = fetch_zip(CALL_PRODUCT, period_str)
        mmddyyyy = period_dt.strftime("%m%d%Y")
        for key, sched in (("ri", "RI"), ("rc", "RC")):
            if key not in targets:
                continue
            raw = zf.read(f"FFIEC CDR Call Schedule {sched} {mmddyyyy}.txt")
            rows, dropped = parse_schedule(raw, _table_columns(names[key]), period_iso, date_text)
            _report(names[key], rows, "IDRSSD", dropped)
            plan.append((names[key], rows))

    if "ubpr" in targets:
        zf, _ = fetch_zip(UBPR_PRODUCT, period_dt.year)
        raw = zf.read(f"FFIEC CDR UBPR Ratios Summary Ratios {period_dt.year}.txt")
        rows, dropped = parse_ubpr(raw, _table_columns("raw_UBPR"), period_dt, period_iso)
        for r in rows:
            r["Reporting Period"] = f"{period_dt.month}/{period_dt.day}/{period_dt.year} 11:59:59 PM"
        _report("raw_UBPR", rows, "ID RSSD", dropped)
        plan.append(("raw_UBPR", rows))

    if dry_run:
        print("Dry run -- nothing written.")
        return
    for table, rows in plan:
        sent = insert(table, rows, schema="raw", batch_size=500)
        print(f"Inserted {sent} rows into raw.{table} ({period_iso}).")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--period", required=True, help="quarter-end as MM/DD/YYYY, e.g. 06/30/2026")
    ap.add_argument("--only", nargs="+", choices=["ri", "rc", "ubpr"], default=None,
                     help="subset of tables (default: all three)")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()
    load(args.period, args.only, args.dry_run)


if __name__ == "__main__":
    sys.exit(main())
