"""Ingests FFIEC's "UBPR Stats -- Four Periods" bulk product into
raw.raw_ubpr_peer_stats + raw.ref_ubpr_stats_fields.

This is genuinely different from the existing raw.raw_UBPR: that table is
bank-level data (one row per institution per period). This one is
peer-group AGGREGATE data (one row per peer-group per period) -- the
"real peer benchmarks" a84 (UBPR Peer-Benchmarking, handed off from
Session 3 2026-10-02) was missing. Verified empirically 2026-10-02: a
sample pull of "Summary Ratios" showed ~2 rows per peer group (one per
reporting period), not one per institution, with a distinct UBPS* field
prefix (vs UBPR* for the bank-level product) and peer-group codes/
descriptions (e.g. "All Bankers Banks in California", asset-size cohorts
1-7/101-104).

See ffiec_bulk_portal.py for how the download itself works (no static
URL, a 3-step ASP.NET postback dance) -- shared with
ffiec_ubpr_rank_ingest.py, the per-institution companion product.

The zip splits into 24 category files (Balance Sheet, Income Statement,
Capital Analysis, Liquidity, etc. -- mirrors how the bank-level UBPR
product is split). Each file has 3 header rows before data starts: field
codes, then long names, then descriptions -- captured here into
ref_ubpr_stats_fields so the mapping/wiring step (Session 3, a84 steps
3-4) doesn't need to go back to FFIEC's own docs. Field codes repeat
across files (the same ratio shown for context in multiple reports), so
raw_ubpr_peer_stats keys on (reporting_period, peer_group, source_file,
field_code), not just the first three.

Usage:
  python -m ingestion.ffiec_ubpr_peer_stats_ingest                # latest available year (auto-detected)
  python -m ingestion.ffiec_ubpr_peer_stats_ingest --year 2026
"""

import argparse
import csv
import io
import sys

from ingestion.ffiec_bulk_portal import core_field_codes, fetch_zip as _portal_fetch_zip
from ingestion.ffiec_bulk_portal import parse_period
from ingestion.supabase_client import upsert

PRODUCT_VALUE = "PerformanceReportingSeriesStats"  # "UBPR Stats -- Four Periods"
WANTED_CODES = core_field_codes("UBPS")


def fetch_zip(requested_year=None):
    return _portal_fetch_zip(PRODUCT_VALUE, requested_year)


def parse_member(fname, raw_bytes):
    text = raw_bytes.decode("utf-8", errors="replace")
    rows = list(csv.reader(io.StringIO(text), delimiter="\t"))
    if len(rows) < 4:
        print(f"  {fname}: fewer than 4 rows, skipping")
        return [], []

    codes, long_names, descriptions = rows[0], rows[1], rows[2]
    field_cols = [i for i in range(3, len(codes)) if codes[i].strip() in WANTED_CODES]
    if not field_cols:
        return [], []

    field_rows = []
    for i in field_cols:
        code = codes[i].strip()
        if not code:
            continue
        field_rows.append({
            "field_code": code,
            "long_name": (long_names[i].strip() if i < len(long_names) else None) or None,
            "description": (descriptions[i].strip() if i < len(descriptions) else None) or None,
            "source_file": fname,
        })

    data_rows = []
    for rec in rows[3:]:
        if len(rec) < 3 or not rec[0].strip():
            continue
        period = parse_period(rec[0])
        peer_group_desc = rec[1].strip()
        peer_group = rec[2].strip()
        for i in field_cols:
            if i >= len(rec):
                continue
            val = rec[i].strip()
            if val == "":
                continue
            try:
                fval = float(val)
            except ValueError:
                continue
            data_rows.append({
                "reporting_period": period.isoformat(),
                "peer_group": peer_group,
                "peer_group_description": peer_group_desc,
                "source_file": fname,
                "field_code": codes[i].strip(),
                "field_value": fval,
            })
    return data_rows, field_rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, default=None,
                     help="calendar year to pull (e.g. 2026); defaults to the latest available")
    args = ap.parse_args()

    zf, label = fetch_zip(args.year)

    all_data, seen_field_codes = [], {}
    for name in zf.namelist():
        if name.strip().lower() == "readme.txt":
            continue
        data_rows, field_rows = parse_member(name, zf.read(name))
        print(f"  {name}: {len(data_rows)} data rows, {len(field_rows)} field codes")
        all_data.extend(data_rows)
        for fr in field_rows:
            seen_field_codes[fr["field_code"]] = fr  # last file wins if a code repeats w/ a different label

    sent_data = upsert(
        "raw_ubpr_peer_stats", all_data,
        on_conflict="reporting_period,peer_group,source_file,field_code",
        schema="raw", batch_size=5000,
    )
    sent_fields = upsert(
        "ref_ubpr_stats_fields", list(seen_field_codes.values()),
        on_conflict="field_code", schema="raw", batch_size=2000,
    )
    print(f"Upserted {sent_data} rows into raw.raw_ubpr_peer_stats ({label}).")
    print(f"Upserted {sent_fields} rows into raw.ref_ubpr_stats_fields.")


if __name__ == "__main__":
    sys.exit(main())
