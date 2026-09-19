"""Ingests Zillow ZHVI (zip-level home value index) into raw.raw_zhvi.

Source: Zillow Research public CSVs (https://www.zillow.com/research/data/),
feed "Zip_zhvi_uc_sfrcondo_tier_0.33_0.67_sm_sa_month.csv" — all-homes,
smoothed/seasonally-adjusted, mid-tier (35th-65th percentile). This is the
feed already implied by raw.raw_zhvi's schema (zip/state/city/metro/
county_name/region_id/size_rank/region_type + period/zhvi) and consumed by
analytics.refresh_branch_opportunity_base's ZHVI-YoY step.

The file is wide (one column per month back to 2000); this script melts it
to raw.raw_zhvi's long format and upserts, so re-runs are idempotent.

Usage:
  python -m ingestion.zhvi_ingest                  # incremental: only periods newer than MAX(period) in raw.raw_zhvi
  python -m ingestion.zhvi_ingest --since 2026-01-31
  python -m ingestion.zhvi_ingest --full-refresh    # reload entire history back to 2000 (slow, ~1M+ rows)
"""

import argparse
import csv
import io
import sys
from datetime import date

import requests

from ingestion.supabase_client import get, upsert

ZHVI_URL = "https://files.zillowstatic.com/research/public_csvs/zhvi/Zip_zhvi_uc_sfrcondo_tier_0.33_0.67_sm_sa_month.csv"

COLUMN_MAP = {
    "RegionID": "region_id",
    "SizeRank": "size_rank",
    "RegionName": "zip",
    "RegionType": "region_type",
    "StateName": "state_name",
    "State": "state",
    "City": "city",
    "Metro": "metro",
    "CountyName": "county_name",
}


def current_max_period():
    rows = get("raw_zhvi", schema="raw", params="select=period&order=period.desc&limit=1")
    if not rows:
        return None
    return date.fromisoformat(rows[0]["period"])


def fetch_csv_text():
    print(f"Downloading {ZHVI_URL} ...")
    r = requests.get(ZHVI_URL, timeout=180)
    r.raise_for_status()
    print(f"  {len(r.content) / 1e6:.1f} MB downloaded")
    return r.text


def melt(csv_text, since):
    reader = csv.DictReader(io.StringIO(csv_text))
    fieldnames = reader.fieldnames
    date_cols = [c for c in fieldnames if c not in COLUMN_MAP]
    keep_cols = [c for c in date_cols if (since is None or date.fromisoformat(c) > since)]
    if not keep_cols:
        return [], date_cols[-1] if date_cols else None

    rows = []
    for rec in reader:
        base = {dst: rec.get(src) or None for src, dst in COLUMN_MAP.items()}
        if base.get("size_rank"):
            base["size_rank"] = int(base["size_rank"])
        for col in keep_cols:
            val = rec.get(col)
            if val in (None, ""):
                continue
            row = dict(base)
            row["period"] = col
            row["zhvi"] = float(val)
            rows.append(row)
    return rows, date_cols[-1]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--since", type=date.fromisoformat, default=None,
                     help="only ingest periods strictly after this date (YYYY-MM-DD)")
    ap.add_argument("--full-refresh", action="store_true",
                     help="ignore current max(period) in the DB and reload full history")
    args = ap.parse_args()

    since = args.since
    if since is None and not args.full_refresh:
        since = current_max_period()
        print(f"Incremental mode: current max(period) in raw.raw_zhvi = {since}")
    elif args.full_refresh:
        print("Full-refresh mode: reloading entire history")

    csv_text = fetch_csv_text()
    rows, latest_col = melt(csv_text, since)

    if not rows:
        print(f"Nothing new to ingest — latest period in the source file is {latest_col}, "
              f"already covered by raw.raw_zhvi.")
        return

    periods = sorted({r["period"] for r in rows})
    print(f"Melted {len(rows)} rows across {len(periods)} period(s): {periods[0]} .. {periods[-1]}")

    sent = upsert("raw_zhvi", rows, on_conflict="zip,period", schema="raw")
    print(f"Upserted {sent} rows into raw.raw_zhvi.")


if __name__ == "__main__":
    sys.exit(main())
