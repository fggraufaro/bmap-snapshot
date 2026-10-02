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

Source: same public portal as the existing raw_UBPR bulk load
(https://cdr.ffiec.gov/public/PWS/DownloadBulkData.aspx), no account or
API key -- it's a plain ASP.NET WebForms page with no login. The page has
no static download URL; every pull does a GET (to scrape a fresh
__VIEWSTATE/__VIEWSTATEGENERATOR -- these are page-load-specific and
can't be hardcoded) followed by a POST of the form fields, which returns
the zip directly as the response body.

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
import zipfile
from datetime import datetime

import requests
from lxml import html as lxml_html

from ingestion.supabase_client import upsert

PORTAL_URL = "https://cdr.ffiec.gov/public/PWS/DownloadBulkData.aspx"
PRODUCT_VALUE = "PerformanceReportingSeriesStats"  # "UBPR Stats -- Four Periods"

# FFIEC's WAF returns 403 for the default python-requests User-Agent --
# confirmed 2026-10-02, a plain browser-like UA is sufficient (no deeper
# bot-detection/CAPTCHA on this bulk-download page, unlike ffiec.gov's
# individual-lookup report pages).
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"}


def _viewstate_fields(doc):
    def field(name):
        els = doc.xpath(f'//input[@name="{name}"]')
        return els[0].get("value", "") if els else ""
    return field("__VIEWSTATE"), field("__VIEWSTATEGENERATOR")


def _scrape_form_state(session):
    # Step 1: plain GET. The product list (ListBox1) and date dropdown
    # (DatesDropDownList) are linked via an ASP.NET autopostback --
    # DatesDropDownList is present but EMPTY until a product is selected
    # (confirmed 2026-10-02), so a fresh GET alone never has year options.
    r = session.get(PORTAL_URL, headers=HEADERS, timeout=30)
    r.raise_for_status()
    doc = lxml_html.fromstring(r.text)
    viewstate, viewstategen = _viewstate_fields(doc)
    if not viewstate:
        raise RuntimeError("Could not find __VIEWSTATE on the FFIEC download page -- page structure may have changed.")

    # Step 2: replicate the product listbox's onchange postback
    # (__doPostBack('ctl00$MainContentHolder$ListBox1','')) to populate
    # DatesDropDownList for the UBPR Stats product specifically.
    data = {
        "__EVENTTARGET": "ctl00$MainContentHolder$ListBox1",
        "__EVENTARGUMENT": "",
        "__LASTFOCUS": "",
        "__VIEWSTATE": viewstate,
        "__VIEWSTATEGENERATOR": viewstategen,
        "ctl00$MainContentHolder$ListBox1": PRODUCT_VALUE,
    }
    r2 = session.post(PORTAL_URL, data=data, headers=HEADERS, timeout=30)
    r2.raise_for_status()
    doc2 = lxml_html.fromstring(r2.text)
    viewstate2, viewstategen2 = _viewstate_fields(doc2)

    years = doc2.xpath('//select[@name="ctl00$MainContentHolder$DatesDropDownList"]/option')
    year_options = [(o.get("value"), o.text_content().strip()) for o in years]
    if not year_options:
        raise RuntimeError("DatesDropDownList still empty after selecting the product -- FFIEC page structure may have changed.")
    return viewstate2, viewstategen2, year_options


def _pick_year_value(year_options, requested_year):
    if requested_year is None:
        return year_options[0]  # newest listed first, verified 2026-10-02
    for value, label in year_options:
        if label.strip() == str(requested_year):
            return value, label
    available = ", ".join(label for _, label in year_options)
    raise RuntimeError(f"Year {requested_year} not available. Options: {available}")


def fetch_zip(requested_year=None):
    session = requests.Session()
    viewstate, viewstategen, year_options = _scrape_form_state(session)
    value, label = _pick_year_value(year_options, requested_year)
    print(f"Pulling UBPR Stats for reporting year {label} (dropdown value {value})...")

    data = {
        "__LASTFOCUS": "",
        "__EVENTTARGET": "",
        "__EVENTARGUMENT": "",
        "__VIEWSTATE": viewstate,
        "__VIEWSTATEGENERATOR": viewstategen,
        "ctl00$MainContentHolder$ListBox1": PRODUCT_VALUE,
        "ctl00$MainContentHolder$DatesDropDownList": value,
        "ctl00$MainContentHolder$FormatType": "TSVRadioButton",
        "ctl00$MainContentHolder$TabStrip1$Download_0": "Download",
    }
    r = session.post(PORTAL_URL, data=data, headers=HEADERS, timeout=120)
    r.raise_for_status()
    if r.headers.get("Content-Type", "").startswith("text/html"):
        raise RuntimeError("FFIEC returned an HTML page instead of a zip -- form fields likely stale/rejected.")
    print(f"  {len(r.content) / 1e6:.1f} MB downloaded")
    return zipfile.ZipFile(io.BytesIO(r.content)), label


def _parse_period(text):
    # "3/31/2026 11:59:59 PM" -> date(2026, 3, 31). Stored as a real DATE
    # column, not text -- a text CYCLE_DATE column elsewhere in this
    # pipeline caused a lexicographic-max bug (see CLAUDE.md), not repeating that here.
    return datetime.strptime(text.strip(), "%m/%d/%Y %I:%M:%S %p").date()


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
        period = _parse_period(rec[0])
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

    all_data, all_fields, seen_field_codes = [], [], {}
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
