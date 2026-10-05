"""a29 step 2: read-only backtest of UBPR NC98 (a Texas-Ratio-style distress
ratio) against the existing noncurrent measure. Definitions and thresholds
were written into the roadmap record (actions/a29) BEFORE any outcome was
computed -- do not change them here to chase a result.

Reads raw.raw_UBPR through PostgREST. Writes nothing to the database.

  python -m analysis.a29_texas_ratio_backtest
"""

import json
import sys

import numpy as np
import pandas as pd
import requests

from ingestion import supabase_client as sc

PERIODS = ["2024-03-31", "2024-06-30", "2024-09-30", "2024-12-31", "2025-03-31",
           "2025-06-30", "2025-09-30", "2025-12-31", "2026-03-31", "2026-06-30"]
Q = {p: i + 1 for i, p in enumerate(PERIODS)}
COLS = {"UBPRNC98": "nc98", "UBPR7414": "c7414", "UBPRE549": "e549", "UBPRD486": "lev", "UBPRE013": "roa"}

# Pre-registered thresholds
AQ_T, CAP_T, EARN_T = 3.0, 6.0, 0.0
FLAGS = {"NC98>=10": ("nc98", 10.0), "NC98>=20": ("nc98", 20.0),
         "existing as coded (7414>0.02)": ("c7414", 0.02), "existing corrected (7414>2.0)": ("c7414", 2.0)}
RNG = np.random.default_rng(1)
N_BOOT = 1000


def load():
    sel = "select=%22ID%20RSSD%22,period," + ",".join(COLS)
    rows, off = [], 0
    while True:
        h = sc._headers("raw")
        h["Range-Unit"] = "items"
        h["Range"] = f"{off}-{off + 999}"
        r = requests.get(f"{sc.SUPA_URL}/rest/v1/raw_UBPR?{sel}&order=period,%22ID%20RSSD%22", headers=h, timeout=120)
        r.raise_for_status()
        chunk = r.json()
        rows += chunk
        if len(chunk) < 1000:
            break
        off += 1000
    df = pd.DataFrame(rows).rename(columns={"ID RSSD": "rssd", **COLS})
    for c in COLS.values():
        df[c] = pd.to_numeric(df[c].replace("", np.nan), errors="coerce")
    df["q"] = df["period"].map(Q)
    return df


def auc(score, y):
    """Mann-Whitney AUC with average ranks for ties; nan if one class is empty."""
    y = np.asarray(y, bool)
    n1, n0 = y.sum(), (~y).sum()
    if n1 == 0 or n0 == 0:
        return np.nan
    ranks = pd.Series(score).rank(method="average").to_numpy()
    return (ranks[y].sum() - n1 * (n1 + 1) / 2) / (n1 * n0)


def build_pairs(df, h, origins):
    """One row per bank-origin with predictors at t and outcome states at t and t+h."""
    a = df[df.q.isin(origins)].copy()
    b = df.copy()
    b["q"] = b["q"] - h
    b = b[["rssd", "q", "e549", "lev", "roa"]].rename(columns={"e549": "e549_f", "lev": "lev_f", "roa": "roa_f"})
    m = a.merge(b, on=["rssd", "q"], how="inner")
    m = m.dropna(subset=["nc98", "c7414", "e549", "lev", "roa", "e549_f", "lev_f", "roa_f"])
    m["aq0"], m["aq1"] = m.e549 >= AQ_T, m.e549_f >= AQ_T
    m["cap0"], m["cap1"] = m.lev < CAP_T, m.lev_f < CAP_T
    m["earn0"], m["earn1"] = m.roa < EARN_T, m.roa_f < EARN_T
    m["any0"] = m.aq0 | m.cap0 | m.earn0
    m["any1"] = m.aq1 | m.cap1 | m.earn1
    return m


def outcome_sets(m):
    return {
        "O_ANY": (m[~m.any0], (~m.any0[~m.any0]) & m.any1[~m.any0]),
        "O_AQ": (m[~m.aq0], m.aq1[~m.aq0]),
        "O_CAP": (m[~m.cap0], m.cap1[~m.cap0]),
        "O_EARN": (m[~m.earn0], m.earn1[~m.earn0]),
    }


def boot_diff(sub, y, n=N_BOOT):
    """Bank-clustered bootstrap of AUC(nc98) - AUC(c7414), plus each AUC's CI."""
    sub = sub.reset_index(drop=True)
    y = np.asarray(y, bool)
    groups = sub.groupby("rssd").indices
    keys = np.array(list(groups.keys()))
    idx_lists = [groups[k] for k in keys]
    d1, d2, dd = [], [], []
    for _ in range(n):
        pick = RNG.integers(0, len(keys), len(keys))
        idx = np.concatenate([idx_lists[i] for i in pick])
        a1 = auc(sub.nc98.to_numpy()[idx], y[idx])
        a2 = auc(sub.c7414.to_numpy()[idx], y[idx])
        if np.isnan(a1) or np.isnan(a2):
            continue
        d1.append(a1); d2.append(a2); dd.append(a1 - a2)
    ci = lambda v: (round(float(np.percentile(v, 2.5)), 3), round(float(np.percentile(v, 97.5)), 3))
    return ci(d1), ci(d2), ci(dd)


def report(label, m, boot=True):
    print(f"\n=== {label}: bank-origin pairs eligible = {len(m)}, banks = {m.rssd.nunique()}")
    out = {}
    for name, (sub, y) in outcome_sets(m).items():
        y = np.asarray(y, bool)
        ev, evb = int(y.sum()), int(sub.loc[y, 'rssd'].nunique()) if y.sum() else 0
        a1, a2 = auc(sub.nc98.to_numpy(), y), auc(sub.c7414.to_numpy(), y)
        line = f"{name}: at-risk pairs {len(sub)}, onset events {ev} (distinct banks {evb}), AUC NC98 {a1:.3f} | AUC 7414 {a2:.3f} | diff {a1 - a2:+.3f}"
        res = {"n": len(sub), "events": ev, "event_banks": evb, "auc_nc98": a1, "auc_7414": a2}
        if boot and ev >= 5:
            c1, c2, cd = boot_diff(sub, y)
            line += f" | 95% CI NC98 {c1}, 7414 {c2}, diff {cd}"
            res.update({"ci_nc98": c1, "ci_7414": c2, "ci_diff": cd})
        print(line)
        out[name] = res
        for fname, (col, thr) in FLAGS.items():
            fl = (sub[col] >= thr) if "NC98" in fname else (sub[col] > thr)
            tp = int((fl & y).sum()); fp = int((fl & ~y).sum())
            prec = tp / (tp + fp) if tp + fp else np.nan
            rec = tp / ev if ev else np.nan
            print(f"    flag {fname:32s} flagged {fl.mean() * 100:5.1f}% of at-risk | precision {prec:.3f} | recall {rec:.3f} | TP {tp} FP {fp}")
            out[name][fname] = {"flagged_share": float(fl.mean()), "precision": prec, "recall": rec, "tp": tp, "fp": fp}
    return out


def main():
    df = load()
    print("rows loaded:", len(df), "| per quarter:", df.groupby("q").size().to_dict())
    results = {}

    m4 = build_pairs(df, 4, range(1, 7))
    results["h4_pooled"] = report("PRIMARY h=4, origins Q1..Q6 pooled", m4)
    print("\nPer-origin AUC for O_ANY (h=4):")
    for t in range(1, 7):
        mt = m4[m4.q == t]
        sub, y = outcome_sets(mt)["O_ANY"]
        y = np.asarray(y, bool)
        print(f"  origin {PERIODS[t - 1]}: at-risk {len(sub)}, events {int(y.sum())}, AUC NC98 {auc(sub.nc98.to_numpy(), y):.3f} | 7414 {auc(sub.c7414.to_numpy(), y):.3f}")

    m2 = build_pairs(df, 2, range(1, 9))
    results["h2_pooled"] = report("SECONDARY h=2, origins Q1..Q8 pooled", m2, boot=True)

    # survivorship: banks present at t (origins Q1..Q6) but absent at t+4
    a = df[df.q.isin(range(1, 7))][["rssd", "q", "nc98", "c7414"]]
    f = df[["rssd", "q"]].assign(q=lambda x: x.q - 4).assign(present=True)
    s = a.merge(f, on=["rssd", "q"], how="left")
    gone, stay = s[s.present.isna()], s[s.present.notna()]
    print(f"\nSurvivorship (h=4): bank-origin pairs absent at t+4 = {len(gone)} of {len(s)}; distinct banks {gone.rssd.nunique()}")
    print(f"  NC98 median: absent {gone.nc98.median():.2f} vs present {stay.nc98.median():.2f}; "
          f"share NC98>=20: absent {100 * (gone.nc98 >= 20).mean():.1f}% vs present {100 * (stay.nc98 >= 20).mean():.1f}%")
    results["survivorship"] = {"absent_pairs": int(len(gone)), "pairs": int(len(s))}
    json.dump(results, open("a29_backtest_results.json", "w"), default=float, indent=1)


if __name__ == "__main__":
    sys.exit(main())
