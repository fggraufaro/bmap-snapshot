"""Refresh loaders for the datasets Session 2 originally loaded by hand (no loader code was ever committed).

  raw.raw_cbp_totals               Census County Business Patterns, ZIP totals   (feeds vw_smb_index_by_zip -> opportunity base)
  raw.raw_business_formation_state Census Business Formation Statistics          (feeds opportunity base)
  raw.raw_irs_migration_state      IRS SOI county migration, summed to state     (feeds opportunity base)
  raw.raw_qcew_state               BLS QCEW annual state wages / employment      (feeds the persona brief)
  raw.raw_cfpb_complaints_trend    CFPB complaints, trailing 90 days per week    (feeds Growth Map)

Each of the first four is a single-vintage replace: the opportunity-base function joins these tables by
state with no vintage filter, so a table holding two vintages would duplicate branch rows. Every loader:
  - builds the new rows from the source,
  - compares them with the live table and prints the differences,
  - applies only when ALLOW_DATASET_LOAD=yes (separate from ALLOW_ANALYTICS_SWAP, which gates the swaps),
    otherwise it stops after the report (a dry run changes nothing),
  - when applying: backup to backup.<table>_<ts> (RLS on), one transaction, row-count check before COMMIT,
  - refuses to apply if the new table is more than 10% smaller than the live one.

Checked against the live tables on 2026-10-06 (reproducing what Session 2 stored): CBP 2023 reproduces
raw_cbp_totals exactly (34,954 rows, 0 differences); IRS 2021-22 reproduces raw_irs_migration_state at exactly
half the stored people/returns/AGI (the 2x double-count in roadmap a25); QCEW 2024 reproduces raw_qcew_state
exactly; Business Formation matches within small Census revisions. The CFPB rule for insufficient_volume
(fewer than 10 complaints in 2024) was inferred from the stored rows, which it separates exactly. The CFPB API
refused the machine this was written on (403), so that loader has not been run against the live API.
"""

import csv
import io
import os
import time
import zipfile
from collections import defaultdict
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal

import requests
from psycopg2 import sql
from psycopg2.extras import execute_values

from ingestion.quarterly_refresh import _backup, _ident, _log, _scalar, _session

GUARD_ENV = "ALLOW_DATASET_LOAD"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
MAX_SHRINK = 0.10

# 50 states + DC
FIPS_TO_ABBR = {
    1: "AL", 2: "AK", 4: "AZ", 5: "AR", 6: "CA", 8: "CO", 9: "CT", 10: "DE", 11: "DC", 12: "FL", 13: "GA", 15: "HI",
    16: "ID", 17: "IL", 18: "IN", 19: "IA", 20: "KS", 21: "KY", 22: "LA", 23: "ME", 24: "MD", 25: "MA", 26: "MI",
    27: "MN", 28: "MS", 29: "MO", 30: "MT", 31: "NE", 32: "NV", 33: "NH", 34: "NJ", 35: "NM", 36: "NY", 37: "NC",
    38: "ND", 39: "OH", 40: "OK", 41: "OR", 42: "PA", 44: "RI", 45: "SC", 46: "SD", 47: "TN", 48: "TX", 49: "UT",
    50: "VT", 51: "VA", 53: "WA", 54: "WV", 55: "WI", 56: "WY",
}
STATES = set(FIPS_TO_ABBR.values())


def _get(url, tries=4, **kw):
    last = None
    for i in range(tries):
        try:
            r = requests.get(url, headers={"User-Agent": UA}, timeout=120, **kw)
            return r
        except requests.RequestException as e:
            last = e
            time.sleep(2 * (i + 1))
    raise RuntimeError(f"could not reach {url}: {last}")


def _norm(v):
    if v is None:
        return None
    if isinstance(v, (Decimal, float)):
        d = Decimal(str(v)).normalize()
        return format(d, "f")
    if isinstance(v, (date, datetime)):
        return v.isoformat()
    return str(v) if not isinstance(v, (int, bool)) else v


# ── fetchers: each returns (rows as list of dicts, one-line description of the vintage) ──────────────
def _fetch_cbp(cur, schema):
    this_year = date.today().year
    for y in range(this_year, this_year - 4, -1):
        r = _get(f"https://www2.census.gov/programs-surveys/cbp/datasets/{y}/zbp{str(y)[2:]}totals.zip")
        if r.status_code == 200 and r.content[:2] == b"PK":
            z = zipfile.ZipFile(io.BytesIO(r.content))
            raw = z.read(z.namelist()[0])
            try:
                text = raw.decode("utf-8")
            except UnicodeDecodeError:
                text = raw.decode("latin-1")
            rows = [{k: (v if v != "" else None) for k, v in rec.items()} for rec in csv.DictReader(io.StringIO(text))]  # blanks load as NULL, as stored
            return rows, f"CBP {y} ZIP totals ({len(rows)} ZIPs)"
    raise RuntimeError("no CBP ZIP-totals file found for the last four years")


def _fetch_bfs(cur, schema):
    r = _get("https://www.census.gov/econ/bfs/csv/bfs_monthly.csv")
    r.raise_for_status()
    months = ("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
    ba = defaultdict(dict)
    for rec in csv.DictReader(io.StringIO(r.text)):
        if rec["series"] == "BA_BA" and rec["sa"] == "U" and rec["naics_sector"] == "TOTAL" and rec["geo"] in STATES:
            vals = [rec[m] for m in months]
            if all(v != "" for v in vals):
                ba[rec["geo"]][int(rec["year"])] = sum(int(v) for v in vals)
    complete = sorted({y for s in ba.values() for y in s})
    years = complete[-3:]
    if years != [2023, 2024, 2025]:
        raise RuntimeError(f"the newest three complete years are {years}, but raw_business_formation_state has fixed columns "
                           "ba2023/ba2024/ba2025; add columns for the new year (and decide how raw tables roll) before loading")
    rows = []
    for st in sorted(ba):
        a, b, c = (ba[st].get(y) for y in years)
        if None in (a, b, c):
            raise RuntimeError(f"{st}: missing a complete year in {years}")
        rows.append({"state": st, "ba2023": a, "ba2024": b, "ba2025": c,
                     "yoy_pct_latest": Decimal(c) / Decimal(b) * 100 - 100})
    for r_ in rows:
        r_["yoy_pct_latest"] = round(r_["yoy_pct_latest"], 2)
    return rows, f"Census BFS business applications {years[0]}-{years[-1]} ({len(rows)} states)"


def _irs_vintage(cur, schema):
    forced = os.environ.get("IRS_VINTAGE")  # e.g. 2022-2023
    cur.execute(sql.SQL("select distinct tax_year_pair from {}").format(_ident((schema, "raw_irs_migration_state"))))
    live = [r[0] for r in cur.fetchall()]
    pair = forced or (live[0] if len(live) == 1 else None)
    if not pair:
        raise RuntimeError(f"raw_irs_migration_state holds {len(live)} vintages {live}; it must hold exactly one")
    return pair, live


def _fetch_irs(cur, schema):
    pair, live = _irs_vintage(cur, schema)
    y1, y2 = pair.split("-")
    code = y1[2:] + y2[2:]
    nxt = f"{int(y1) + 1}-{int(y2) + 1}"
    nr = requests.head(f"https://www.irs.gov/pub/irs-soi/countyinflow{nxt[2:4]}{nxt[7:9]}.csv", headers={"User-Agent": UA}, timeout=30)
    _log(f"  IRS vintage {pair} (the one currently loaded{' - forced by IRS_VINTAGE' if os.environ.get('IRS_VINTAGE') else ''}); "
         f"newer vintage {nxt} is {'AVAILABLE (set IRS_VINTAGE to move; a new series, scoring policy call)' if nr.status_code == 200 else 'not published'}")
    files = {}
    for kind in ("inflow", "outflow"):
        r = _get(f"https://www.irs.gov/pub/irs-soi/county{kind}{code}.csv")
        if r.status_code != 200 or len(r.content) < 100000:
            raise RuntimeError(f"IRS county{kind}{code}.csv not available (HTTP {r.status_code})")
        files[kind] = list(csv.DictReader(io.StringIO(r.content.decode("latin-1"))))
    agg = defaultdict(lambda: [0] * 6)
    # Total Migration-US rows only (code 97 / county 0). Summing the 97 sub-rows as well was the 2x bug (roadmap a25).
    for rec in files["inflow"]:
        if rec["y1_statefips"] == "97" and rec["y1_countyfips"] == "0":
            a = agg[int(rec["y2_statefips"])]
            a[0] += int(rec["n1"]); a[1] += int(rec["n2"]); a[2] += int(rec["agi"])
    for rec in files["outflow"]:
        if rec["y2_statefips"] == "97" and rec["y2_countyfips"] == "0":
            a = agg[int(rec["y1_statefips"])]
            a[3] += int(rec["n1"]); a[4] += int(rec["n2"]); a[5] += int(rec["agi"])
    rows = []
    for fips, ab in sorted(FIPS_TO_ABBR.items(), key=lambda kv: kv[1]):
        if fips not in agg:
            raise RuntimeError(f"{ab}: no IRS rows")
        ri, pi, ai, ro, po, ao = agg[fips]
        rows.append({"state": ab, "tax_year_pair": pair, "returns_in": ri, "people_in": pi, "agi_in_thousands": ai,
                     "returns_out": ro, "people_out": po, "agi_out_thousands": ao,
                     "net_people": pi - po, "net_agi_thousands": ai - ao})
    return rows, f"IRS SOI {pair} county files summed to state ({len(rows)} states)"


def _fetch_qcew(cur, schema):
    this_year = date.today().year
    for y in range(this_year, this_year - 4, -1):
        r = _get(f"https://data.bls.gov/cew/data/api/{y}/a/industry/10.csv")
        if r.status_code == 200 and len(r.content) > 100000:
            break
    else:
        raise RuntimeError("no QCEW annual file found for the last four years")
    rows = []
    for rec in csv.DictReader(io.StringIO(r.text)):
        if rec["own_code"] == "0" and rec["agglvl_code"] == "50" and rec["area_fips"].endswith("000"):
            fips = int(rec["area_fips"][:2])
            if fips in FIPS_TO_ABBR:
                rows.append({"state": FIPS_TO_ABBR[fips], "year": int(rec["year"]),
                             "avg_estabs": Decimal(rec["annual_avg_estabs"]), "avg_emplvl": Decimal(rec["annual_avg_emplvl"]),
                             "avg_annual_pay": Decimal(rec["avg_annual_pay"]), "avg_wkly_wage": Decimal(rec["annual_avg_wkly_wage"]),
                             "pay_yoy_chg": Decimal(rec["oty_avg_annual_pay_chg"]),
                             "pay_yoy_pct_chg": Decimal(rec["oty_avg_annual_pay_pct_chg"]),
                             "emplvl_yoy_pct_chg": Decimal(rec["oty_annual_avg_emplvl_pct_chg"])})
    rows.sort(key=lambda x: x["state"])
    return rows, f"BLS QCEW {y} annual, all industries ({len(rows)} states)"


CFPB_URL = "https://www.consumerfinance.gov/data-research/consumer-complaints/search/api/v1/"


def _cfpb_count(company, start, end):
    r = None
    for i in range(4):
        r = requests.get(CFPB_URL, params={"company": company, "size": 0, "date_received_min": start.isoformat(),
                                           "date_received_max": end.isoformat(), "no_aggs": "true"},
                         headers={"User-Agent": UA, "Accept": "application/json"}, timeout=60)
        if r.status_code in (429, 500, 502, 503, 504):
            time.sleep(3 * (i + 1))
            continue
        break
    if r.status_code == 403:
        raise RuntimeError("the CFPB API refused the request (HTTP 403). It blocks some networks; run this from Railway "
                           "or check the User-Agent. Nothing was changed.")
    r.raise_for_status()
    total = r.json()["hits"]["total"]
    return total["value"] if isinstance(total, dict) else total


def _fetch_cfpb_trend(cur, schema, today=None, count=_cfpb_count):
    """Two rows per institution: the latest completed week (Thursday) and the week before it, each the
    complaints received in the 90 days ending that Thursday. The volume flag follows the stored rule:
    fewer than 10 complaints in 2024."""
    today = today or date.today()
    w = today - timedelta(days=1)
    while w.weekday() != 3:  # Thursday, as in the stored weeks (2026-09-10, 2026-09-17)
        w -= timedelta(days=1)
    cur.execute(sql.SQL("select inst_key, cfpb_company_name, complaints_2024 from {} order by inst_key")
                .format(_ident((schema, "raw_cfpb_complaints"))))
    insts = cur.fetchall()
    rows = []
    for k, name, c24 in insts:
        for wk in (w - timedelta(days=7), w):
            n = count(name, wk - timedelta(days=89), wk)
            rows.append({"inst_key": k, "week_ending": wk, "trailing_90d_count": n, "insufficient_volume": c24 < 10})
        time.sleep(0.15)
    return rows, f"CFPB trailing-90-day counts for weeks ending {w - timedelta(days=7)} and {w} ({len(insts)} institutions)"


# ── dataset specs ────────────────────────────────────────────────────────────────────────────────────
SPECS = {
    "cbp": dict(table="raw_cbp_totals", key=["zip"], mode="replace", loaded_at=False, fetch=_fetch_cbp,
                cols=["zip", "name", "emp_nf", "emp", "qp1_nf", "qp1", "ap_nf", "ap", "est", "city", "stabbr", "cty_name"]),
    "business_formation": dict(table="raw_business_formation_state", key=["state"], mode="replace", loaded_at=True,
                               fetch=_fetch_bfs, cols=["state", "ba2023", "ba2024", "ba2025", "yoy_pct_latest"]),
    "irs_migration": dict(table="raw_irs_migration_state", key=["state"], mode="replace", loaded_at=True, fetch=_fetch_irs,
                          cols=["state", "tax_year_pair", "returns_in", "people_in", "agi_in_thousands", "returns_out",
                                "people_out", "agi_out_thousands", "net_people", "net_agi_thousands"]),
    "qcew": dict(table="raw_qcew_state", key=["state"], mode="replace", loaded_at=True, fetch=_fetch_qcew,
                 cols=["state", "year", "avg_estabs", "avg_emplvl", "avg_annual_pay", "avg_wkly_wage", "pay_yoy_chg",
                       "pay_yoy_pct_chg", "emplvl_yoy_pct_chg"]),
    "cfpb_trend": dict(table="raw_cfpb_complaints_trend", key=["inst_key", "week_ending"], mode="upsert", loaded_at=True,
                       fetch=_fetch_cfpb_trend, cols=["inst_key", "week_ending", "trailing_90d_count", "insufficient_volume"]),
}


def _live_rows(cur, schema, spec):
    cur.execute(sql.SQL("select {} from {}").format(sql.SQL(", ").join(map(sql.Identifier, spec["cols"])),
                                                    _ident((schema, spec["table"]))))
    return [dict(zip(spec["cols"], r)) for r in cur.fetchall()]


def _compare(spec, live, new):
    kf = lambda r: tuple(_norm(r[c]) for c in spec["key"])
    L, N = {kf(r): r for r in live}, {kf(r): r for r in new}
    if len(N) != len(new):
        raise RuntimeError("the new data has duplicate keys; refusing")
    changed = [(k, L[k], N[k]) for k in N if k in L and any(_norm(L[k][c]) != _norm(N[k][c]) for c in spec["cols"])]
    return dict(live=len(L), new=len(N), added=len(set(N) - set(L)), removed=len(set(L) - set(N)),
                changed=len(changed), same=len(set(N) & set(L)) - len(changed), samples=changed[:4])


def _report(name, spec, diff, desc):
    _log(f"  source: {desc}")
    _log(f"  live {diff['live']} rows -> new {diff['new']} rows: {diff['same']} unchanged, {diff['changed']} changed, "
         f"{diff['added']} added, {diff['removed']} removed (upsert keeps removed rows)" if spec["mode"] == "upsert" else
         f"  live {diff['live']} rows -> new {diff['new']} rows: {diff['same']} unchanged, {diff['changed']} changed, "
         f"{diff['added']} added, {diff['removed']} removed")
    for k, old, new in diff["samples"]:
        delta = {c: (old[c], new[c]) for c in spec["cols"] if _norm(old[c]) != _norm(new[c])}
        _log(f"    e.g. {k}: {delta}")


def _extra_checks(name, live, new):
    if name == "irs_migration" and live:
        L = {r["state"]: r for r in live}
        worst = 0.0
        halved = 0
        for r in new:
            o = L.get(r["state"])
            if not o:
                continue
            rate = lambda x: x["net_agi_thousands"] / (x["agi_in_thousands"] + x["agi_out_thousands"])
            worst = max(worst, abs(rate(o) - rate(r)))
            halved += o["people_in"] == 2 * r["people_in"]
        _log(f"  IRS check: {halved} of {len(new)} states have stored people_in exactly 2x the new value (the a25 double count); "
             f"largest change in the scoring rate (net AGI / gross AGI) = {worst:.2e}")


def refresh(name, schema="raw", backup_schema="backup", apply=None, fetch=None):
    spec = dict(SPECS[name])
    if fetch:
        spec["fetch"] = fetch
    apply = (os.environ.get(GUARD_ENV) == "yes") if apply is None else apply
    _log(f"[{name}] {'APPLY' if apply else 'DRY RUN (set ' + GUARD_ENV + '=yes to apply)'}")
    with _session("refresh_" + name) as conn:
        cur = conn.cursor()
        new, desc = spec["fetch"](cur, schema)
        if not new:
            raise RuntimeError("the source returned no rows; nothing to load")
        for r in new:
            missing = [c for c in spec["cols"] if c not in r]
            if missing:
                raise RuntimeError(f"fetched rows are missing columns {missing}")
        live = _live_rows(cur, schema, spec)
        diff = _compare(spec, live, new)
        _report(name, spec, diff, desc)
        _extra_checks(name, live, new)
        if spec["mode"] == "replace" and diff["live"] and diff["new"] < diff["live"] * (1 - MAX_SHRINK):
            raise RuntimeError(f"new data has {diff['new']} rows vs {diff['live']} live (more than {int(MAX_SHRINK * 100)}% smaller); refusing")
        if not apply:
            _log("  dry run only: nothing was changed.")
            return diff
        target = (schema, spec["table"])
        _backup(cur, target, backup_schema, spec["table"], time.strftime("%Y%m%d_%H%M%S"))
        cols = spec["cols"] + (["loaded_at"] if spec["loaded_at"] else [])
        values = [tuple(r[c] for c in spec["cols"]) + ((datetime.now(timezone.utc),) if spec["loaded_at"] else ()) for r in new]
        col_sql = sql.SQL(", ").join(map(sql.Identifier, cols))
        conn.autocommit = False
        try:
            if spec["mode"] == "replace":
                cur.execute(sql.SQL("DELETE FROM {}").format(_ident(target)))
                execute_values(cur, sql.SQL("INSERT INTO {} ({}) VALUES %s").format(_ident(target), col_sql), values)
                n = _scalar(cur, sql.SQL("select count(*) from {}").format(_ident(target)))
                if n != len(new):
                    raise RuntimeError(f"after the reload the table has {n} rows, expected {len(new)}; rolled back")
            else:
                upd = sql.SQL(", ").join(sql.SQL("{0} = EXCLUDED.{0}").format(sql.Identifier(c)) for c in cols if c not in spec["key"])
                execute_values(cur, sql.SQL("INSERT INTO {} ({}) VALUES %s ON CONFLICT ({}) DO UPDATE SET {}")
                               .format(_ident(target), col_sql, sql.SQL(", ").join(map(sql.Identifier, spec["key"])), upd), values)
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.autocommit = True
        _log(f"  applied: {len(new)} rows written to {schema}.{spec['table']}.")
        return diff


def refresh_cbp():
    refresh("cbp")


def refresh_business_formation():
    refresh("business_formation")


def refresh_irs_migration():
    refresh("irs_migration")


def refresh_qcew():
    refresh("qcew")


def refresh_cfpb_trend():
    refresh("cfpb_trend")
