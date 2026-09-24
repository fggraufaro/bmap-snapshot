"""Ingests NCUA quarterly Call Report data into raw."Raw_cu_fs220" and
raw.raw_cu_branches.

Source: NCUA's public quarterly Call Report bulk download
(https://ncua.gov/analysis/credit-union-corporate-call-report-data/quarterly-data),
a static, no-auth ZIP per quarter at a predictable URL:
  https://ncua.gov/files/publications/analysis/call-report-data-YYYY-MM.zip
containing (among other schedules) FS220.txt and "Credit Union Branch
Information.txt", whose headers are a 1:1 match for raw."Raw_cu_fs220" and
raw.raw_cu_branches respectively (verified against a real 2026-06 download —
see CU_NUMBER/CYCLE_DATE/JOIN_NUMBER/UPDATE_DATE/ACCT_* for FS220 and
CU_NUMBER/CYCLE_DATE/JOIN_NUMBER/SiteId/... for branch info — so rows pass
through as-is, no field mapping needed).

This is the credit-union half of a14. The bank half (FFIEC Call Report
schedules RC/RI/RCE/RIE + UBPR) is NOT covered here — FFIEC has no
equivalent static bulk file; its only programmatic path is the CDR
"Public Data Distribution" SOAP webservice, which requires a registered
CDR account (free, but a new credential like a17's Census key) and
retrieves data per-institution (RetrieveFacsimile takes a single fiID),
not as one bulk file. See SESSION_COORDINATION.md / sprint notes for
that follow-up.

Usage:
  python -m ingestion.ncua_fs220_ingest                  # latest available quarter (auto-detected)
  python -m ingestion.ncua_fs220_ingest --period 2026-06  # explicit quarter (YYYY-MM, quarter-end month)
  python -m ingestion.ncua_fs220_ingest --only fs220      # skip branch info
  python -m ingestion.ncua_fs220_ingest --only branches   # skip FS220
"""

import argparse
import csv
import io
import sys
import zipfile
from datetime import date

import requests

from ingestion.supabase_client import upsert

NCUA_URL_TMPL = "https://ncua.gov/files/publications/analysis/call-report-data-{year}-{month:02d}.zip"
QUARTER_END_MONTHS = (3, 6, 9, 12)

FS220_ENTRY = "FS220.txt"
BRANCHES_ENTRY = "Credit Union Branch Information.txt"


def _quarter_candidates(start=None):
    """Yields (year, month) quarter-end pairs, most recent first, walking
    backward from the most recent quarter-end on or before `start`
    (defaults to today)."""
    today = start or date.today()
    y, qm = today.year, max((q for q in QUARTER_END_MONTHS if q <= today.month), default=None)
    if qm is None:
        y, qm = y - 1, QUARTER_END_MONTHS[-1]
    while True:
        yield y, qm
        idx = QUARTER_END_MONTHS.index(qm)
        y, qm = (y - 1, QUARTER_END_MONTHS[-1]) if idx == 0 else (y, QUARTER_END_MONTHS[idx - 1])


def find_latest_available(max_tries=6):
    for y, m in _quarter_candidates():
        if max_tries <= 0:
            break
        max_tries -= 1
        url = NCUA_URL_TMPL.format(year=y, month=m)
        r = requests.head(url, timeout=30, allow_redirects=True)
        if r.status_code == 200:
            return y, m
        print(f"  {url} -> {r.status_code}, trying prior quarter")
    raise RuntimeError("No NCUA quarterly file found in the last few quarters")


def fetch_zip(year, month):
    url = NCUA_URL_TMPL.format(year=year, month=month)
    print(f"Downloading {url} ...")
    r = requests.get(url, timeout=180)
    r.raise_for_status()
    print(f"  {len(r.content) / 1e6:.1f} MB downloaded")
    return zipfile.ZipFile(io.BytesIO(r.content))


def parse_rows(csv_text):
    reader = csv.DictReader(io.StringIO(csv_text))
    rows = []
    for rec in reader:
        row = {k: (v if v not in (None, "") else None) for k, v in rec.items()}
        rows.append(row)
    return rows


def ingest_fs220(z):
    with z.open(FS220_ENTRY) as f:
        csv_text = f.read().decode("utf-8", errors="replace")
    rows = parse_rows(csv_text)
    if not rows:
        print("No rows parsed from FS220.txt — skipping.")
        return
    cycle_dates = sorted({r.get("CYCLE_DATE") for r in rows})
    print(f"Parsed {len(rows)} FS220 rows, CYCLE_DATE(s): {cycle_dates}")
    sent = upsert("Raw_cu_fs220", rows, on_conflict="CU_NUMBER,CYCLE_DATE", schema="raw", batch_size=1000)
    print(f'Upserted {sent} rows into raw."Raw_cu_fs220".')


def ingest_branches(z):
    with z.open(BRANCHES_ENTRY) as f:
        csv_text = f.read().decode("utf-8", errors="replace")
    rows = parse_rows(csv_text)
    if not rows:
        print("No rows parsed from 'Credit Union Branch Information.txt' — skipping.")
        return
    cycle_dates = sorted({r.get("CYCLE_DATE") for r in rows})
    print(f"Parsed {len(rows)} branch rows, CYCLE_DATE(s): {cycle_dates}")
    sent = upsert(
        "raw_cu_branches", rows,
        on_conflict="CU_NUMBER,CYCLE_DATE,JOIN_NUMBER,SiteId",
        schema="raw", batch_size=1000,
    )
    print(f"Upserted {sent} rows into raw.raw_cu_branches.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--period", type=str, default=None,
                     help="quarter-end as YYYY-MM (e.g. 2026-06); defaults to the latest available quarter")
    ap.add_argument("--only", choices=["fs220", "branches"], default=None,
                     help="ingest only one of the two files (default: both)")
    args = ap.parse_args()

    if args.period:
        y, m = (int(x) for x in args.period.split("-"))
        if m not in QUARTER_END_MONTHS:
            print(f"WARNING: {m:02d} is not a standard quarter-end month {QUARTER_END_MONTHS}")
    else:
        print("Auto-detecting latest available NCUA quarter...")
        y, m = find_latest_available()

    z = fetch_zip(y, m)

    if args.only != "branches":
        ingest_fs220(z)
    if args.only != "fs220":
        ingest_branches(z)


if __name__ == "__main__":
    sys.exit(main())
