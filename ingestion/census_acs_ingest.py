"""Ingests Census ACS 5-Year estimates (household income + population by
age, at zip-code/ZCTA level) into raw.raw_income and raw.raw_population.

BLOCKED on a Census API key (a17) — as of 2026-09-11, api.census.gov
returns "Missing Key" for every data query, including trivial/single-row
ones (this evidently got stricter at some point; smaller unauthenticated
requests used to work). Get a free key at
https://api.census.gov/data/key_signup.html and set CENSUS_API_KEY.
Metadata endpoints (variable/group definitions) don't need a key, which
is how the table/variable mapping below was verified without one.

Tables used (matched against raw_income/raw_population's existing columns,
which still carry a prior manual export's human-relabeled names):
  - S1903 (Median Income in the Past 12 Months, by age of householder)
      income          = S1903_C03_001E  (Total)
      income u25      = S1903_C03_011E  (15 to 24 years - the only "under
                         25" bracket S1903 has; confirmed via variable
                         metadata, matches the existing 4-bucket shape)
      income 25 to 44 = S1903_C03_012E
      income 45 to 64 = S1903_C03_013E
      income over 65  = S1903_C03_014E
  - S0101 (Age and Sex), 5-year age brackets summed into raw_population's
    6 buckets (verified additive against S0101_C01_001E "Total" using the
    variable definitions - NOT yet checked against real ZCTA-level output,
    since that needs the key this script doesn't have yet):
      total           = S0101_C01_001E
      total_under_15  = 002E + 003E + 004E   (Under 5, 5-9, 10-14)
      total_15to25    = 005E + 006E          (15-19, 20-24)
      total_25to35    = 007E + 008E          (25-29, 30-34)
      total_35to45    = 009E + 010E          (35-39, 40-44)
      total_45to65    = 011E + 012E + 013E + 014E  (45-49...60-64)
      total_over65    = S0101_C01_030E       (65 years and over, direct)

Both tables are published as one ACS 5-year vintage per year (~every
December, ~12 months after the vintage's end year) at ZCTA level
nationally in a single request — no pagination needed (~33.8k ZCTAs).

Usage:
  python -m ingestion.census_acs_ingest                  # latest likely-available vintage (heuristic, see below)
  python -m ingestion.census_acs_ingest --year 2024
"""

import argparse
import os
import sys
from datetime import date

import requests

from ingestion.supabase_client import upsert

CENSUS_API_KEY = os.environ.get("CENSUS_API_KEY", "")

INCOME_VARS = {
    "income": "S1903_C03_001E",
    "income u25": "S1903_C03_011E",
    "income 25 to 44": "S1903_C03_012E",
    "income 45 to 64": "S1903_C03_013E",
    "income over 65": "S1903_C03_014E",
}
POP_TOTAL_VAR = "S0101_C01_001E"
POP_BUCKET_VARS = {
    "total_under_15": ["S0101_C01_002E", "S0101_C01_003E", "S0101_C01_004E"],
    "total_15to25": ["S0101_C01_005E", "S0101_C01_006E"],
    "total_25to35": ["S0101_C01_007E", "S0101_C01_008E"],
    "total_35to45": ["S0101_C01_009E", "S0101_C01_010E"],
    "total_45to65": ["S0101_C01_011E", "S0101_C01_012E", "S0101_C01_013E", "S0101_C01_014E"],
    "total_over65": ["S0101_C01_030E"],
}


def latest_likely_vintage():
    """ACS 5-year vintage Y is typically released in December of Y+1.
    Heuristic only - override with --year once the real release date for
    a given year is known."""
    today = date.today()
    return today.year - 2 if today.month < 12 else today.year - 1


def _get(url, params):
    if not CENSUS_API_KEY:
        raise RuntimeError(
            "CENSUS_API_KEY is not set. api.census.gov requires a key for data queries "
            "(confirmed 2026-09-11 - even trivial unauthenticated requests are rejected "
            "with 'Missing Key'). Get one at https://api.census.gov/data/key_signup.html "
            "(this is action a17)."
        )
    params = dict(params, key=CENSUS_API_KEY)
    r = requests.get(url, params=params, timeout=120)
    r.raise_for_status()
    return r.json()


def fetch_income(year):
    url = f"https://api.census.gov/data/{year}/acs/acs5/subject"
    get_vars = ["NAME", "GEO_ID"] + list(INCOME_VARS.values())
    payload = _get(url, {"get": ",".join(get_vars), "for": "zip code tabulation area:*"})
    header, rows = payload[0], payload[1:]
    idx = {name: header.index(name) for name in get_vars}

    out = []
    for r in rows:
        row = {
            "Geography": r[idx["GEO_ID"]],
            "Geographic Area Name": r[idx["NAME"]],
            "YEAR": str(year),
        }
        for col, var in INCOME_VARS.items():
            v = r[idx[var]]
            row[col] = None if v in (None, "", "-666666666") else str(v)
        out.append(row)
    return out


def fetch_population(year):
    url = f"https://api.census.gov/data/{year}/acs/acs5/subject"
    all_vars = sorted({POP_TOTAL_VAR} | {v for vs in POP_BUCKET_VARS.values() for v in vs})
    get_vars = ["NAME", "GEO_ID"] + all_vars
    payload = _get(url, {"get": ",".join(get_vars), "for": "zip code tabulation area:*"})
    header, rows = payload[0], payload[1:]
    idx = {name: header.index(name) for name in get_vars}

    def num(r, var):
        v = r[idx[var]]
        if v in (None, "", "-666666666"):
            return None
        return int(v)

    out = []
    for r in rows:
        row = {
            "Geography": r[idx["GEO_ID"]],
            "Geographic Area Name": r[idx["NAME"]],
            "YEAR": str(year),
            "total": num(r, POP_TOTAL_VAR),
        }
        for bucket, vars_ in POP_BUCKET_VARS.items():
            parts = [num(r, v) for v in vars_]
            row[bucket] = sum(p for p in parts if p is not None) if any(p is not None for p in parts) else None
        out.append(row)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, default=None,
                     help="ACS 5-year vintage (ending year), e.g. 2024; defaults to a release-date heuristic")
    args = ap.parse_args()
    year = args.year or latest_likely_vintage()
    print(f"Fetching Census ACS 5-year vintage {year} (income: S1903, population: S0101)...")

    income_rows = fetch_income(year)
    print(f"  income: {len(income_rows)} ZCTAs")
    pop_rows = fetch_population(year)
    print(f"  population: {len(pop_rows)} ZCTAs")

    sent_income = upsert("raw_income", income_rows, on_conflict="Geography,YEAR", schema="raw", batch_size=5000)
    print(f"Upserted {sent_income} rows into raw.raw_income.")

    sent_pop = upsert("raw_population", pop_rows, on_conflict="Geography,YEAR", schema="raw", batch_size=5000)
    print(f"Upserted {sent_pop} rows into raw.raw_population.")


if __name__ == "__main__":
    sys.exit(main())
