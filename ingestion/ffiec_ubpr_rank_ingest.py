"""Ingests FFIEC's "UBPR Rank -- Four Periods" bulk product into
raw.raw_ubpr_rank + raw.ref_ubpr_rank_fields.

Companion to ffiec_ubpr_peer_stats_ingest.py's peer-group AGGREGATES
(a84, UBPR Peer-Benchmarking): this product is each institution's own
PERCENTILE RANK (0-100) within its peer group, per ratio -- e.g. "this
bank's loan-to-deposit ratio is in the 63rd percentile vs its peers."
Verified empirically 2026-10-02: header is "Reporting Period / ID RSSD /
Peer Group / <field codes>" (vs Stats' "Reporting Period / Peer Group
Description / Peer Group / <field codes>") -- institution-level, not
peer-group-level. An institution can belong to more than one peer group
at once (its asset-size cohort plus a specialty/geographic one), so
peer_group is part of the key, not just metadata.

See ffiec_bulk_portal.py for how the download itself works (no static
URL, a 3-step ASP.NET postback dance), shared with
ffiec_ubpr_peer_stats_ingest.py.

Usage:
  python -m ingestion.ffiec_ubpr_rank_ingest                # latest available year (auto-detected)
  python -m ingestion.ffiec_ubpr_rank_ingest --year 2026
"""

import argparse
import csv
import io
import sys

from ingestion.ffiec_bulk_portal import fetch_zip as _portal_fetch_zip
from ingestion.ffiec_bulk_portal import parse_period
from ingestion.supabase_client import upsert

PRODUCT_VALUE = "PerformanceReportingSeriesRank"  # "UBPR Rank -- Four Periods"


def fetch_zip(requested_year=None):
    return _portal_fetch_zip(PRODUCT_VALUE, requested_year)


def parse_member(fname, raw_bytes):
    text = raw_bytes.decode("utf-8", errors="replace")
    rows = list(csv.reader(io.StringIO(text), delimiter="\t"))
    if len(rows) < 4:
        print(f"  {fname}: fewer than 4 rows, skipping")
        return [], []

    codes, long_names, descriptions = rows[0], rows[1], rows[2]
    field_cols = list(range(3, len(codes)))

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
        id_rssd = rec[1].strip()
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
                "id_rssd": id_rssd,
                "peer_group": peer_group,
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

    # Processed and upserted one file at a time, not accumulated across all
    # 26 files first -- this product melts to ~30M rows (vs Stats' ~0.5M),
    # too much to hold as Python dicts in memory at once, and per-file
    # upserts mean a mid-run failure only costs the current file's
    # progress (re-running main() is safe either way -- upsert is
    # idempotent -- but this avoids redoing everything already landed).
    total_sent, seen_field_codes = 0, {}
    members = [n for n in zf.namelist() if n.strip().lower() != "readme.txt"]
    for idx, name in enumerate(members, 1):
        data_rows, field_rows = parse_member(name, zf.read(name))
        for fr in field_rows:
            seen_field_codes[fr["field_code"]] = fr
        sent = upsert(
            "raw_ubpr_rank", data_rows,
            on_conflict="reporting_period,id_rssd,peer_group,source_file,field_code",
            schema="raw", batch_size=5000,
        )
        total_sent += sent
        print(f"  [{idx}/{len(members)}] {name}: {len(data_rows)} data rows -> upserted {sent}")

    sent_fields = upsert(
        "ref_ubpr_rank_fields", list(seen_field_codes.values()),
        on_conflict="field_code", schema="raw", batch_size=2000,
    )
    print(f"Upserted {total_sent} total rows into raw.raw_ubpr_rank ({label}).")
    print(f"Upserted {sent_fields} rows into raw.ref_ubpr_rank_fields.")


if __name__ == "__main__":
    sys.exit(main())
