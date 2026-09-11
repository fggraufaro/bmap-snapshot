"""Ingests NCUA quarterly Call Report data (FS220 schedule) into raw."Raw_cu_fs220".

Source: NCUA's public quarterly Call Report bulk download
(https://ncua.gov/analysis/credit-union-corporate-call-report-data/quarterly-data),
a static, no-auth ZIP per quarter at a predictable URL:
  https://ncua.gov/files/publications/analysis/call-report-data-YYYY-MM.zip
containing (among other schedules) FS220.txt, whose header is a 1:1 match
for raw."Raw_cu_fs220"'s columns (verified against a real download — see
CU_NUMBER/CYCLE_DATE/JOIN_NUMBER/UPDATE_DATE/ACCT_* — so rows pass through
as-is, no field mapping needed).

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
  python -m ingestion.ncua_fs220_ingest --period 2026-03  # explicit quarter (YYYY-MM, quarter-end month)
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


def fetch_fs220(year, month):
    url = NCUA_URL_TMPL.format(year=year, month=month)
    print(f"Downloading {url} ...")
    r = requests.get(url, timeout=180)
    r.raise_for_status()
    print(f"  {len(r.content) / 1e6:.1f} MB downloaded")
    z = zipfile.ZipFile(io.BytesIO(r.content))
    with z.open("FS220.txt") as f:
        return f.read().decode("utf-8", errors="replace")


def parse_rows(csv_text):
    reader = csv.DictReader(io.StringIO(csv_text))
    rows = []
    for rec in reader:
        row = {k: (v if v not in (None, "") else None) for k, v in rec.items()}
        rows.append(row)
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--period", type=str, default=None,
                     help="quarter-end as YYYY-MM (e.g. 2026-03); defaults to the latest available quarter")
    args = ap.parse_args()

    if args.period:
        y, m = (int(x) for x in args.period.split("-"))
        if m not in QUARTER_END_MONTHS:
            print(f"WARNING: {m:02d} is not a standard quarter-end month {QUARTER_END_MONTHS}")
    else:
        print("Auto-detecting latest available NCUA quarter...")
        y, m = find_latest_available()

    csv_text = fetch_fs220(y, m)
    rows = parse_rows(csv_text)
    if not rows:
        print("No rows parsed from FS220.txt — aborting without writing.")
        return

    cycle_dates = sorted({r.get("CYCLE_DATE") for r in rows})
    print(f"Parsed {len(rows)} FS220 rows, CYCLE_DATE(s): {cycle_dates}")

    sent = upsert("Raw_cu_fs220", rows, on_conflict="CU_NUMBER,CYCLE_DATE", schema="raw", batch_size=1000)
    print(f'Upserted {sent} rows into raw."Raw_cu_fs220".')


if __name__ == "__main__":
    sys.exit(main())
