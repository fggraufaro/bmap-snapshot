"""Ingests FDIC Summary of Deposits (SOD, annual, as-of June 30) into raw.raw_sod.

Source: FDIC's public BankFind Suite API (https://api.fdic.gov/banks/sod),
no key required. SOD is published once a year (~September for the prior
June 30 snapshot) — this script auto-detects the latest year with data.

Field names from the API already match raw_sod's existing (quoted,
uppercase) column names 1:1 — verified against a live pull. Only the
subset of columns raw_sod actually has is kept; anything else the API
returns (e.g. its own row "ID") is dropped rather than sent and rejected
by PostgREST.

CITYBR (branch city) vs CITY (HQ city) — see CLAUDE.md's ghost-data
gotcha. This script doesn't need to choose between them; it loads both,
as-is, same as the rest of raw_sod's columns.

Usage:
  python -m ingestion.fdic_sod_ingest                  # latest available year (auto-detected)
  python -m ingestion.fdic_sod_ingest --year 2025
"""

import argparse
import sys
from datetime import date

import requests

from ingestion.supabase_client import upsert

FDIC_API = "https://api.fdic.gov/banks/sod"
PAGE_SIZE = 10000

# The subset of the API's response fields that raw.raw_sod actually has
# columns for (verified against information_schema.columns).
RAW_SOD_COLUMNS = {
    "YEAR", "CERT", "NAMEFULL", "ADDRESBR", "BRNUM", "UNINUMBR", "STALPBR",
    "ADDRESS", "ASSET", "BKCLASS", "CALL", "CHARTER", "CHRTAGNN", "CHRTAGNT",
    "CITY", "CLCODE", "CNTRYNA", "DENOVO", "DEPDOM", "DEPSUM", "DOCKET",
    "ESCROW", "FDICDBS", "FDICNAME", "FED", "FEDNAME", "INSAGNT1", "INSBRDD",
    "INSBRTS", "INSURED", "OCCDIST", "OCCNAME", "REGAGNT", "RSSDID",
    "SPECDESC", "SPECGRP", "STALP", "STCNTY", "STNAME", "UNIT", "ZIP",
    "BKMO", "BRCENM", "BRSERTYP", "CBSA_DIV_NAMB", "CITY2BR", "CITYBR",
    "CNTRYNAB", "CNTYNAMB", "CNTYNUMB", "CONSOLD", "CSABR", "CSANAMBR",
    "DEPSUMBR", "DIVISIONB", "METROBR", "MICROBR", "MSABR", "MSANAMB",
    "NAMEBR", "NECNAMB", "NECTABR", "PLACENUM", "SIMS_ACQUIRED_DATE",
    "SIMS_DESCRIPTION", "SIMS_ESTABLISHED_DATE", "SIMS_LATITUDE",
    "SIMS_LONGITUDE", "SIMS_PROJECTION", "STCNTYBR", "STNAMEBR", "STNUMBR",
    "USA", "ZIPBR", "CITYHCR", "HCTMULT", "NAMEHCR", "RSSDHCR", "STALPHCR",
}
NUMERIC_COLUMNS = {"ASSET", "FED", "DEPSUMBR", "SIMS_LATITUDE", "SIMS_LONGITUDE"}


def _year_has_data(year):
    r = requests.get(FDIC_API, params={"filters": f"YEAR:{year}", "limit": 1}, timeout=30)
    r.raise_for_status()
    return r.json()["meta"]["total"] > 0


def find_latest_available_year():
    year = date.today().year
    for y in range(year, year - 4, -1):
        if _year_has_data(y):
            return y
    raise RuntimeError("No FDIC SOD data found in the last few years")


def fetch_year(year):
    rows = []
    offset = 0
    while True:
        r = requests.get(FDIC_API, params={"filters": f"YEAR:{year}", "limit": PAGE_SIZE, "offset": offset},
                          timeout=60)
        r.raise_for_status()
        payload = r.json()
        batch = [d["data"] for d in payload["data"]]
        rows.extend(batch)
        total = payload["meta"]["total"]
        print(f"  fetched {len(rows)}/{total}")
        if len(rows) >= total or not batch:
            break
        offset += PAGE_SIZE
    return rows


def to_row(rec):
    row = {}
    for col in RAW_SOD_COLUMNS:
        val = rec.get(col)
        if val is None:
            row[col] = None
        elif col in NUMERIC_COLUMNS:
            row[col] = val
        else:
            row[col] = str(val)
    return row


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, default=None,
                     help="SOD year to ingest (as-of June 30); defaults to the latest year with published data")
    args = ap.parse_args()

    year = args.year
    if year is None:
        print("Auto-detecting latest available FDIC SOD year...")
        year = find_latest_available_year()
    print(f"Fetching FDIC SOD for YEAR={year} ...")

    records = fetch_year(year)
    if not records:
        print(f"No records returned for YEAR={year} — aborting without writing.")
        return

    rows = [to_row(r) for r in records]
    print(f"Parsed {len(rows)} branch rows for YEAR={year}.")

    sent = upsert("raw_sod", rows, on_conflict="UNINUMBR,YEAR", schema="raw", batch_size=5000)
    print(f"Upserted {sent} rows into raw.raw_sod.")


if __name__ == "__main__":
    sys.exit(main())
