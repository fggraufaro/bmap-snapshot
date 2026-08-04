"""
bmap_assessment_doc.py — Verlocity $10K BMAP Assessment (Word doc generator)
=============================================================================
Produces the paid $10K Assessment deliverable as a branded .docx report.
Full-network coverage (every branch, no top-N slice) — this is the structural
difference vs. the free BMAP Snapshot (bmap_snapshot.py), which only shows
the top 5 branches as an outreach teaser.

Same Supabase source as bmap_snapshot.py / bmap_board_brief.py.
Uses python-docx (matches production Python stack — Flask/Railway).

Usage (CLI):
    python bmap_assessment_doc.py --inst_key bank_463735 --name "Hancock Whitney Bank"

Scope note: this is a first structural draft to test what full-network data
+ AI narrative looks like at scale. Session-based content (the live working
call) is intentionally NOT in this doc — per the 7/22 scope doc, the document
alone is not the $10K product; the session is. See closing section.
"""

import os
import sys
import argparse
import requests
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

try:
    import anthropic
except ImportError:
    anthropic = None

import json

# ── Config — matches current bmap_snapshot.py production pattern ──
# (previous version of this file used a hardcoded legacy anon key, which
# is dead now that RLS + service-role-only access is in place — fixed here)
SUPA_URL = "https://tuiiywphoynbmkxpoyps.supabase.co"
SUPA_KEY = os.environ.get("SUPABASE_SERVICE_KEY", "")
if not SUPA_KEY:
    print("  ⚠  SUPABASE_SERVICE_KEY is not set — Supabase calls will fail with 401.")

ANTH_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
OUT_DIR  = Path(".")

# ── Brand colors — matches Verlocity_Brand_Guidelines_R2.pdf + bmap_snapshot.py exactly ──
def rgb(h): return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

NAVY   = rgb("083D5F")   # Primary Dark Blue
TEAL   = rgb("02A7C2")   # Primary Light Blue
JET    = rgb("213141")   # Jet Black
EMERALD = rgb("66CC99")
LEMON   = rgb("CDD61A")
GRAY3  = rgb("778899")
GRAY_FILL = "F5F5F2"

# Zone colors — identical palette to bmap_snapshot.py, not a separate guess
ZONE_COLOR = {
    "Invest":  "27500A", "Analyze": "185FA5", "Defend": "854F0B", "Justify": "A32D2D",
}
ZONE_LIGHT = {
    "Invest":  "EAF3DE", "Analyze": "E6F1FB", "Defend": "FFF3E0", "Justify": "FCEBEB",
}
ZONE_HEX_MPL = {  # matplotlib wants '#rrggbb'
    z: f"#{h}" for z, h in ZONE_COLOR.items()
}
FONT_HEAD = "Inter"   # falls back to a system serif/sans if not installed locally

# ═══════════════════════════════════════════════════════════════
# DATA FETCH — full network, no truncation
# ═══════════════════════════════════════════════════════════════

# SCHEMA_MAP mirrors bmap_snapshot.py exactly — table→schema routing,
# since PostgREST needs Accept-Profile for anything outside 'public'.
SCHEMA_MAP = {
    "branch_opportunity_base":        "analytics",
    "bank_financial_snapshot_latest": "analytics",
    "branches_master_v2":             "geo",
}

def supabase(table, params):
    url = f"{SUPA_URL}/rest/v1/{table}?{params}"
    schema = SCHEMA_MAP.get(table, "public")
    headers = {"apikey": SUPA_KEY, "Authorization": f"Bearer {SUPA_KEY}"}
    if schema != "public":
        headers["Accept-Profile"] = schema
    r = requests.get(url, headers=headers, timeout=30)
    if r.status_code != 200:
        print(f"  ⚠ Supabase {table} error {r.status_code}: {r.text[:200]}")
        return []
    return r.json()


def supabase_rpc(fn_name, payload, timeout=20):
    """Call a Postgres function via PostgREST's /rpc/ endpoint (e.g. the
    parametrized branches_within_radius(lat, lon, radius, exclude_bank_id))."""
    url = f"{SUPA_URL}/rest/v1/rpc/{fn_name}"
    headers = {"apikey": SUPA_KEY, "Authorization": f"Bearer {SUPA_KEY}",
               "Content-Type": "application/json"}
    r = requests.post(url, headers=headers, json=payload, timeout=timeout)
    if r.status_code != 200:
        print(f"  ⚠ RPC {fn_name} error {r.status_code}: {r.text[:200]}")
        return []
    return r.json()


def fetch_full_network_data(ik):
    """Pull FULL branch network — no limit=N slice. This is the structural
    difference vs. the free Snapshot."""
    print(f"  Fetching full branch network for {ik}...")
    branches = supabase(
        "branch_opportunity_base",
        f"inst_key=eq.{ik}&select=uninumbr,namebr,citybr,stalpbr,latest_dep,"
        "yoy_deposits,opportunity_score,opportunity_zone,matrix_quadrant,"
        "priority_tier,market_growth_score,rel_growth_norm,namefull,"
        "household_income,yoy_income_growth,total_population,yoy_pop_growth,"
        "zhvi_yoy_pct&order=opportunity_score.desc",
    )
    print(f"  ✓ {len(branches)} branches (full network, uncapped)")

    print(f"  Fetching financial snapshot...")
    fin_arr = supabase(
        "bank_financial_snapshot_latest",
        f"inst_key=eq.{ik}&select=roa,nim,efficiency_ratio,dep_yoy_pct,"
        "dep_qoq_pct,cost_of_funds_pct,tier1_capital_pct,net_income_yoy_pct,"
        "total_assets,total_deposits,period",
    )
    fin = fin_arr[0] if fin_arr else {}

    print(f"  Fetching network competitor target...")
    tgt_arr = supabase(
        "vw_network_top_targets",
        f"my_inst_key=eq.{ik}&select=target_institution,branches_in_radius,"
        "avg_vuln_score,avg_yoy_pct,target_roa,target_efficiency_ratio,dominant_zone"
        "&order=network_rank.asc&limit=3",
    )

    print(f"  Fetching branch geography for map...")
    branches_geo = fetch_branch_geo(ik)
    print(f"  ✓ {len(branches_geo)} branches geocoded")

    branch_strategy = fetch_branch_competitive_strategy(ik, branches, branches_geo)

    return {
        "inst_key": ik,
        "branches": branches,
        "fin": fin,
        "targets": tgt_arr,
        "branches_geo": branches_geo,
        "branch_strategy": branch_strategy,
    }


def _vuln_tier(score):
    """vuln_score is an intentionally uncapped composite (base 0-100 x up to
    ~4x stacked risk multipliers) - by design, not a bug. Showing the raw
    number next to a 0-100% financial table reads as broken, so we display
    a qualitative tier instead of implying false precision."""
    s = _sf(score)
    if s >= 150:
        return "Critical"
    if s >= 100:
        return "High"
    if s >= 60:
        return "Elevated"
    return "Moderate"


def _sf(v, default=0.0):
    try:
        return float(v) if v is not None else default
    except (TypeError, ValueError):
        return default


# ═══════════════════════════════════════════════════════════════
# VISUALS — matplotlib charts + geographic map, embedded as PNGs.
# McKinsey-style: clean, minimal chrome, brand colors, direct labels
# instead of legends where possible.
# ═══════════════════════════════════════════════════════════════
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

plt.rcParams["font.family"] = "sans-serif"
plt.rcParams["font.sans-serif"] = ["DejaVu Sans", "Arial", "Inter"]
plt.rcParams["axes.edgecolor"] = "#CCCCCC"
plt.rcParams["axes.linewidth"] = 0.6

NAVY_HEX = "#083D5F"
GRAY_HEX = "#8A8A80"


def chart_zone_distribution(zones, path):
    """Horizontal bar, one bar per zone, real brand zone colors, count labeled directly."""
    order = ["Invest", "Analyze", "Defend", "Justify"]
    vals = [zones.get(z, 0) for z in order]
    colors = [ZONE_HEX_MPL[z] for z in order]

    fig, ax = plt.subplots(figsize=(7.2, 2.3), dpi=200)
    y = range(len(order))
    bars = ax.barh(y, vals, color=colors, height=0.62)
    for i, v in enumerate(vals):
        ax.text(v + max(vals) * 0.02, i, str(v), va="center", fontsize=12,
                fontweight="bold", color="#222222")
    ax.set_yticks(y)
    ax.set_yticklabels(order, fontsize=11, color="#222222")
    ax.invert_yaxis()
    ax.set_xticks([])
    for spine in ["top", "right", "bottom"]:
        ax.spines[spine].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.tick_params(left=False)
    fig.tight_layout(pad=0.6)
    fig.savefig(path, transparent=True)
    plt.close(fig)


def chart_top_bottom_branches(top5, bottom3, path):
    """Diverging horizontal bar — top branches (green, growth) vs bottom (red, decline)."""
    rows = [(b["namebr"], _sf(b["opportunity_score"]), True) for b in top5] + \
           [(b["namebr"], _sf(b["opportunity_score"]), False) for b in bottom3]
    names = [r[0] for r in rows]
    scores = [r[1] for r in rows]
    colors = [ZONE_HEX_MPL["Invest"] if r[2] else ZONE_HEX_MPL["Justify"] for r in rows]

    fig, ax = plt.subplots(figsize=(7.2, 3.2), dpi=200)
    y = range(len(rows))
    ax.barh(y, scores, color=colors, height=0.6)
    for i, v in enumerate(scores):
        ax.text(v + 1.5, i, f"{v:.0f}", va="center", fontsize=10, fontweight="bold", color="#222222")
    ax.set_yticks(y)
    ax.set_yticklabels(names, fontsize=9.5, color="#222222")
    ax.invert_yaxis()
    ax.set_xlim(0, 105)
    ax.set_xticks([])
    ax.axvline(x=50, color="#CCCCCC", linewidth=0.8, linestyle="--")
    for spine in ["top", "right", "bottom"]:
        ax.spines[spine].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.tick_params(left=False)
    fig.tight_layout(pad=0.6)
    fig.savefig(path, transparent=True)
    plt.close(fig)


def chart_financial_benchmark(fin, path):
    """Grouped bar — actual metric vs benchmark, normalized to comparable scale per metric."""
    metrics = [
        ("ROA", _sf(fin.get("roa")), 1.0),
        ("NIM", _sf(fin.get("nim")), 3.0),
        ("Efficiency", _sf(fin.get("efficiency_ratio")), 60.0),
        ("Dep. YoY", _sf(fin.get("dep_yoy_pct")), 2.0),
        ("Tier 1 Cap.", _sf(fin.get("tier1_capital_pct")), 8.0),
    ]
    labels = [m[0] for m in metrics]
    actual = [m[1] for m in metrics]
    bench = [m[2] for m in metrics]

    x = range(len(labels))
    w = 0.32
    fig, ax = plt.subplots(figsize=(7.2, 2.8), dpi=200)
    ax.bar([i - w/2 for i in x], actual, width=w, color=NAVY_HEX, label="Mid Penn")
    ax.bar([i + w/2 for i in x], bench, width=w, color="#C9CED6", label="Benchmark")
    ax.set_xticks(list(x))
    ax.set_xticklabels(labels, fontsize=9.5)
    ax.set_yticks([])
    for spine in ["top", "right", "left"]:
        ax.spines[spine].set_visible(False)
    ax.legend(frameon=False, fontsize=9, loc="upper right")
    fig.tight_layout(pad=0.6)
    fig.savefig(path, transparent=True)
    plt.close(fig)


def chart_branch_map(branches_geo, path):
    """Geographic bubble map — lat/lon scatter, sized by deposits, colored by zone.
    No basemap tiles (no network access to a tile service in this environment) —
    branch clustering and zone concentration are the useful signal here regardless."""
    if not branches_geo:
        return False

    fig, ax = plt.subplots(figsize=(7.2, 5.4), dpi=200)
    for zone in ["Justify", "Defend", "Analyze", "Invest"]:  # draw Invest last (on top)
        pts = [b for b in branches_geo if b.get("opportunity_zone") == zone]
        if not pts:
            continue
        lons = [b["lon"] for b in pts]
        lats = [b["lat"] for b in pts]
        sizes = [max(_sf(b.get("latest_dep")) / 3e6, 18) for b in pts]
        ax.scatter(lons, lats, s=sizes, c=ZONE_HEX_MPL[zone], alpha=0.75,
                   edgecolors="white", linewidths=0.5, label=zone)

    ax.set_xticks([])
    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.set_facecolor("#FAFAF8")
    ax.legend(frameon=False, fontsize=10, loc="lower left", markerscale=0.6)
    ax.set_aspect("equal", adjustable="datalim")
    fig.tight_layout(pad=0.3)
    fig.savefig(path, transparent=False, facecolor="#FAFAF8")
    plt.close(fig)
    return True


def fetch_branch_geo(ik):
    """Lat/lon for the branch map — joined from geo.branches_master_v2.
    branch_id in geo matches uninumbr in branch_opportunity_base."""
    rows = supabase(
        "branch_opportunity_base",
        f"inst_key=eq.{ik}&select=uninumbr,opportunity_zone,latest_dep",
    )
    if not rows:
        return []
    ids = ",".join(str(r["uninumbr"]) for r in rows)
    geo_rows = supabase(
        "branches_master_v2",
        f"branch_id=in.({ids})&select=branch_id,lat,lon",
    )
    geo_rows = geo_rows if isinstance(geo_rows, list) else []
    geo_map = {g["branch_id"]: g for g in geo_rows}
    out = []
    for r in rows:
        g = geo_map.get(r["uninumbr"])
        if g and g.get("lat") and g.get("lon"):
            out.append({**r, "lat": g["lat"], "lon": g["lon"]})
    return out


# ═══════════════════════════════════════════════════════════════
# THE 16 PLAYS — from BMAP_Methodology_Part1.docx (Princeton Partners
# Group proprietary methodology). Zone (row) x Quadrant (col) -> play.
# Q1=Grow&Perform Q2=Invest&Protect Q3=Maintain&Improve Q4=Rationalize&Exit
# ═══════════════════════════════════════════════════════════════
PLAY_MATRIX = {
    ("Invest",  "Q1"): "Aggressive Acquisition",
    ("Invest",  "Q2"): "Market Domination",
    ("Invest",  "Q3"): "Growth Opportunity",
    ("Invest",  "Q4"): "Niche Opportunity",
    ("Analyze", "Q1"): "Competitive Defense",
    ("Analyze", "Q2"): "Grow Share",
    ("Analyze", "Q3"): "Maintain",
    ("Analyze", "Q4"): "Efficiency Review",
    ("Defend",  "Q1"): "Urgent Competitive Push",
    ("Defend",  "Q2"): "Targeted Defense",
    ("Defend",  "Q3"): "Steady State",
    ("Defend",  "Q4"): "Efficiency Review",
    ("Justify", "Q1"): "Exit Strategy",
    ("Justify", "Q2"): "Asset Optimization",
    ("Justify", "Q3"): "Performance Improvement",
    ("Justify", "Q4"): "Rationalize",
}

# Acquisition intensity per play — drives whether a branch gets an
# acquisition-posture writeup or a no-campaign/diagnostic one, per
# "A Rationalize play branch does not receive an acquisition campaign.
# A Grow Share play branch does." (Methodology Part 1)
PLAY_ACQUISITION_POSTURE = {
    "Aggressive Acquisition": "Maximum acquisition budget — all channels appropriate.",
    "Market Domination": "High acquisition budget — dominance posture, deter competitor response.",
    "Growth Opportunity": "Moderate acquisition budget — efficiency-focused channel mix.",
    "Niche Opportunity": "Targeted, modest budget — diagnose the specific niche before briefing media.",
    "Competitive Defense": "Retention-first budget — stabilize before acquisition.",
    "Grow Share": "Selective acquisition budget — test before scaling.",
    "Maintain": "Low acquisition budget — retention and cross-sell focus.",
    "Efficiency Review": "No new acquisition investment until market trajectory is assessed.",
    "Urgent Competitive Push": "Defensive acquisition budget — rate campaigns, switching offers.",
    "Targeted Defense": "Retention-focused budget — loyalty and relationship offers.",
    "Steady State": "Minimal budget — operational efficiency review only.",
    "Exit Strategy": "No acquisition investment — full strategic review required.",
    "Asset Optimization": "No new acquisition — maximize return from existing customer base.",
    "Performance Improvement": "Modest operational investment — no acquisition campaign.",
    "Rationalize": "No investment pending diagnosis — board-level attention for $100M+ branches.",
}


PLAY_MEDIA_BRIEF = {
    "Aggressive Acquisition": "Target: new-to-bank households in the branch's 10-mile radius. Product: CD or HYSA. KPIs: new accounts, deposit volume, cost per new account.",
    "Market Domination": "Target: in-market consumers and competitive switchers. Goal: increase share of wallet and deter competitor response.",
    "Growth Opportunity": "Prioritize high-intent audiences. Avoid broad awareness spend that the market's modest growth rate cannot return.",
    "Niche Opportunity": "Identify the specific audience segment driving the score before briefing media.",
    "Competitive Defense": "Priority: stop outflows. Secondary: selective new account acquisition for high-value segments only.",
    "Grow Share": "Target: deposit-active households in Invest and Analyze zones. Product: CD or savings. Measure cost per new account vs MediaPredict forecast.",
    "Maintain": "Minimize acquisition cost. Prioritize retention of high-balance customers.",
    "Efficiency Review": "No brief generated. Diagnostic phase. Understand market trajectory before any media allocation.",
    "Urgent Competitive Push": "Target: competitor bank customers. Message: switching value proposition. KPI: captured accounts from identifiable competitor customers.",
    "Targeted Defense": "Retention posture. Existing customer focus. Minimize churn.",
    "Steady State": "No acquisition brief. Operational focus only.",
    "Exit Strategy": "No brief generated. Strategic escalation required.",
    "Asset Optimization": "No acquisition brief. Existing customer focus — CD renewals, balance growth, product consolidation.",
    "Performance Improvement": "No brief generated.",
    "Rationalize": "No brief generated. Board-level attention for branches with $100M+ in deposits.",
}


def get_play(zone, matrix_quadrant):
    """Parse 'Q1 - Grow and Perform' -> 'Q1', look up (zone, quadrant) in
    the 16-play matrix. Falls back gracefully if either is missing/unrecognized."""
    if not zone or not matrix_quadrant:
        return None
    q = matrix_quadrant.split("-")[0].split("–")[0].strip().split()[0] if matrix_quadrant else None
    return PLAY_MATRIX.get((zone, q))


def determine_adaptive_radius(density_1mi_count, branch_deposits):
    """
    Density + deposit-size adaptive radius rule (per user directive):
    - Rural OR low-deposit branch -> widen the search. This is the safe
      default: a thin or low-value market shouldn't risk missing the
      few competitors that do exist.
    - Crowded AND high-deposit -> can tighten to 0.5mi. Only be this
      aggressive when BOTH signals confirm it's a real, contestable,
      valuable market -- not on density or deposits alone.
    Real validation: Camden (dense, $113M branch) -> 20 competitors within
    1mi alone. Millersburg (rural, $534M branch -- the HQ) -> 0 competitors
    even at 1mi. Same fixed radius cannot serve both.
    """
    is_rural    = density_1mi_count < 3
    is_low_dep  = branch_deposits < 30_000_000
    is_dense    = density_1mi_count >= 15
    is_high_dep = branch_deposits >= 100_000_000

    if is_rural or is_low_dep:
        return 10.0
    if is_dense and is_high_dep:
        return 0.5
    if is_dense:
        return 1.0
    return 3.0


def fetch_branch_competitive_strategy(ik, branches, branches_geo):
    """Per-branch adaptive-radius competitor lookup via ONE call to
    branches_within_radius_batch(), grouped locally by branch. Previously
    made one branches_within_radius() call per branch (59+ sequential
    round-trips for a mid-size network) which was exceeding the Gunicorn
    worker timeout in production and killing the connection mid-request.
    Same density + deposit rule, same size filter -- only the fetch pattern
    changed."""
    try:
        bank_id = int(ik.replace("bank_", ""))
    except ValueError:
        print(f"  ⚠ Could not derive numeric bank_id from '{ik}' (credit union or "
              f"non-standard inst_key?) — skipping branch-level competitive strategy.")
        return []

    print(f"  Fetching adaptive-radius competitive strategy for {len(branches)} branches "
          f"(single batched call)...")
    all_candidates = supabase_rpc("branches_within_radius_batch", {
        "p_inst_key": ik, "p_exclude_bank_id": bank_id, "p_max_radius_miles": 10.0,
    })
    if not isinstance(all_candidates, list):
        all_candidates = []

    candidates_by_branch = {}
    for c in all_candidates:
        candidates_by_branch.setdefault(c["my_uninumbr"], []).append(c)

    results = []
    for b in branches:
        candidates = candidates_by_branch.get(b["uninumbr"], [])

        density_1mi = sum(1 for c in candidates if _sf(c.get("distance_miles")) <= 1.0)
        deposits = _sf(b.get("latest_dep"))
        radius = determine_adaptive_radius(density_1mi, deposits)
        # Same size filter as branch_target_competitors (Methodology Part 2):
        # competitor deposits between 0.10x and 5x the client branch's own
        # deposits. Without this, giant national-bank hub branches (Wells
        # Fargo, PNC, etc. with $1-8B at a single location) get picked as
        # the "capture opportunity," which is not a realistic local target.
        min_dep, max_dep = deposits * 0.10, deposits * 5.0
        filtered = sorted(
            (c for c in candidates
             if _sf(c.get("distance_miles")) <= radius
             and min_dep <= _sf(c.get("deposits")) <= max_dep),
            key=lambda c: -_sf(c.get("deposits"))
        )
        top_competitor = filtered[0] if filtered else None
        top3_competitors = filtered[:3]

        results.append({
            "namebr": b.get("namebr"),
            "citybr": b.get("citybr"),
            "stalpbr": b.get("stalpbr"),
            "deposits": deposits,
            "radius_mi": radius,
            "tier": ("Dense/High-Value" if radius == 0.5 else
                     "Dense" if radius == 1.0 else
                     "Suburban" if radius == 3.0 else "Low-Density"),
            "competitor_count": len(filtered),
            "top_competitor": top_competitor,
            "top3_competitors": top3_competitors,
        })
    print(f"  ✓ {len(results)} branches assessed")
    return results


DEEP_DIVE_THRESHOLD = 20  # <20 branches -> assess every branch. >=20 -> curate top opportunities.


def build_branch_deep_dives(branches, branch_strategy):
    """Per-branch full assessment: zone, quadrant, play, priority tier, deposits,
    named competitor, and per-branch $ capture scenario.

    <20 branches: every branch gets a full writeup (deep coverage).
    >=20 branches: curated to the branches with the biggest actual $ capture
    opportunity -- ranked by their named competitor's deposits (the real
    contestable dollar figure), not just opportunity_score alone.
    """
    strategy_by_name = {(r["namebr"], r["citybr"], r["stalpbr"]): r for r in branch_strategy}

    enriched = []
    for b in branches:
        key = (b.get("namebr"), b.get("citybr"), b.get("stalpbr"))
        strat = strategy_by_name.get(key)
        play = get_play(b.get("opportunity_zone"), b.get("matrix_quadrant"))
        top_comp = strat.get("top_competitor") if strat else None
        capture_pool = _sf(top_comp.get("deposits")) if top_comp else 0.0
        enriched.append({
            "branch": b,
            "strategy": strat,
            "play": play,
            "capture_pool": capture_pool,
        })

    deep_mode = len(branches) < DEEP_DIVE_THRESHOLD
    if deep_mode:
        selected = enriched  # every branch
    else:
        # Biggest actual $ opportunity first, not just highest score
        selected = sorted(enriched, key=lambda e: -e["capture_pool"])[:15]

    return selected, deep_mode


def summarize_network(d):
    """Aggregate stats for narrative context — NOT a per-branch dump.
    Keeps AI input tokens flat regardless of network size (14 branches or 300)."""
    br = d["branches"]
    n = len(br)
    zones = {"Invest": 0, "Analyze": 0, "Defend": 0, "Justify": 0}
    total_dep = 0.0
    yoy_vals = []
    for b in br:
        z = b.get("opportunity_zone")
        if z in zones:
            zones[z] += 1
        total_dep += _sf(b.get("latest_dep"))
        yoy_vals.append(_sf(b.get("yoy_deposits")))

    avg_yoy = sum(yoy_vals) / len(yoy_vals) if yoy_vals else 0
    avg_score = sum(_sf(b.get("opportunity_score")) for b in br) / n if n else 0

    top5 = sorted(br, key=lambda b: -_sf(b.get("opportunity_score")))[:5]
    bottom3 = sorted(br, key=lambda b: _sf(b.get("opportunity_score")))[:3]

    return {
        "branch_count": n,
        "zones": zones,
        "total_deposits_B": total_dep / 1e9,
        "avg_yoy_pct": avg_yoy * 100,
        "avg_score": avg_score,
        "top5": top5,
        "bottom3": bottom3,
    }


# ═══════════════════════════════════════════════════════════════
# AI NARRATIVE — aggregated context only, not per-branch
# ═══════════════════════════════════════════════════════════════

def summarize_branch_strategy(branch_strategy):
    """Bank-wide roll-up by density/deposit tier - feeds both the narrative
    and the 'as a full bank' strategy view, not just per-branch detail."""
    tiers = {}
    for r in branch_strategy:
        t = r["tier"]
        tiers.setdefault(t, {"count": 0, "deposits": 0.0})
        tiers[t]["count"] += 1
        tiers[t]["deposits"] += r["deposits"]
    named_hits = [r for r in branch_strategy if r.get("top_competitor")]
    return {"tiers": tiers, "named_hits": named_hits}


def get_narratives(bank_name, summary, fin, targets, branch_strategy=None, dives=None):
    branch_strategy = branch_strategy or []
    dives = dives or []
    if not ANTH_KEY or not anthropic:
        print("  ⚠ No ANTHROPIC_API_KEY — using placeholder narratives")
        return _placeholder_narratives(dives)

    zones = summary["zones"]
    top5_str = "; ".join(
        f"{b['namebr']} ({b['citybr']}, {b['stalpbr']}) — score {_sf(b['opportunity_score']):.0f}, "
        f"${_sf(b['latest_dep'])/1e6:.0f}M deposits, {_sf(b['yoy_deposits'])*100:+.1f}% YoY"
        for b in summary["top5"]
    )
    bottom3_str = "; ".join(
        f"{b['namebr']} ({b['citybr']}, {b['stalpbr']}) — score {_sf(b['opportunity_score']):.0f}"
        for b in summary["bottom3"]
    )
    target_str = "; ".join(
        f"{t.get('target_institution','—')} — {t.get('branches_in_radius','—')} branches exposed, "
        f"{_vuln_tier(t.get('avg_vuln_score'))} vulnerability, deposit YoY {_sf(t.get('avg_yoy_pct')):.1f}%"
        for t in targets
    ) or "No qualifying network-level target identified"

    bs_summary = summarize_branch_strategy(branch_strategy)
    tier_str = "; ".join(
        f"{t}: {v['count']} branches, ${v['deposits']/1e6:.0f}M deposits"
        for t, v in bs_summary["tiers"].items()
    )
    # Best 4 named, real, adaptive-radius findings for the model to draw on directly
    named_examples = sorted(bs_summary["named_hits"], key=lambda r: -r["deposits"])[:4]
    named_str = "; ".join(
        f"{r['namebr']} ({r['citybr']}, {r['stalpbr']}, {r['tier']}, {r['radius_mi']}mi radius) — "
        f"largest nearby competitor {r['top_competitor']['bank_name']} "
        f"{r['top_competitor']['distance_miles']:.2f}mi away, "
        f"${_sf(r['top_competitor']['deposits'])/1e6:.0f}M deposits"
        for r in named_examples
    ) or "No adaptive-radius competitor matches found"

    ctx = f"""
Bank: {bank_name}
Full network: {summary['branch_count']} branches, ${summary['total_deposits_B']:.1f}B total deposits
Zone distribution: Invest {zones['Invest']} | Analyze {zones['Analyze']} | Defend {zones['Defend']} | Justify {zones['Justify']}
Network avg opportunity score: {summary['avg_score']:.0f}/100
Network avg deposit YoY: {summary['avg_yoy_pct']:+.1f}%

Top 5 branches by opportunity: {top5_str}
Bottom 3 branches by opportunity: {bottom3_str}

Financial health: ROA {_sf(fin.get('roa')):.2f}% | NIM {_sf(fin.get('nim')):.2f}% | Efficiency {_sf(fin.get('efficiency_ratio')):.1f}%
Deposit YoY {_sf(fin.get('dep_yoy_pct')):+.1f}% | Cost of funds {_sf(fin.get('cost_of_funds_pct')):.2f}% | Tier 1 {_sf(fin.get('tier1_capital_pct')):.1f}%
Net income YoY {_sf(fin.get('net_income_yoy_pct')):+.1f}%

Network-level competitive targets: {target_str}

Branch-level adaptive-radius competitive strategy (radius scaled per branch by local
density + deposit size -- dense/high-value branches get as tight as 0.5mi, rural or
low-deposit branches widen to 10mi, since one fixed radius misses the real picture
for a network spanning both dense suburbs and rural markets):
By tier: {tier_str}
Named examples: {named_str}
"""

    # Per-branch demographic context for the batched audience narrative --
    # ONE call covers all deep-dive branches, not one call per branch, since
    # fetch_branch_competitive_strategy already makes N sequential RPC calls
    # and adding N more AI calls would compound that latency risk.
    deep_dive_ctx = ""
    if dives:
        lines = []
        for e in dives:
            b = e["branch"]
            lines.append(
                f"- {b.get('namebr')} ({b.get('citybr')}, {b.get('stalpbr')}): "
                f"household income ${_sf(b.get('household_income')):.0f}, "
                f"income YoY {_sf(b.get('yoy_income_growth'))*100:+.1f}%, "
                f"population {_sf(b.get('total_population')):.0f}, "
                f"pop YoY {_sf(b.get('yoy_pop_growth'))*100:+.1f}%, "
                f"home value YoY {_sf(b.get('zhvi_yoy_pct')):+.1f}%, "
                f"zone {b.get('opportunity_zone')}, play {e['play'] or 'n/a'}"
            )
        deep_dive_ctx = "\n\nBranches needing a 2-3 sentence audience signal (real Census/ZHVI data, " \
                        "no fabricated personas -- Verlocity's persona layer is still in development):\n" + \
                        "\n".join(lines)

    system = """You are writing the $10K Verlocity BMAP Assessment for a community bank CFO/CEO audience.
Tone: precise, CFO-appropriate. No superlatives, no urgency language, no self-referential commentary.
State facts, name specific branches/competitors, quantify everything possible.
Return ONLY valid JSON, no markdown fences:
{
  "exec_summary": "3-4 sentences. The single most important finding, stated plainly, with a number.",
  "network_narrative": "2-3 sentences on what the zone distribution reveals about the network's overall position.",
  "competitive_narrative": "2-3 sentences naming the specific network-level target and why it is vulnerable.",
  "financial_narrative": "2-3 sentences on what the financial metrics mean together — not a list restated as prose.",
  "capture_strategy_narrative": "3-4 sentences on the branch-level adaptive-radius findings. Name at least one specific dense/high-value branch with its named largest nearby competitor and distance, and contrast the tactical approach that implies (rate/digital competition at close range) against what the low-density branches need instead (defense and wallet-share deepening, since there is often no competitor within the adaptive radius to capture from). This is the 'win deposits by branch AND as a full bank' section.",
  "next_step": "2-3 sentences. A specific, named recommendation tied to the top opportunity branches.",
  "branch_audiences": {"Branch Name (City, ST)": "2-3 sentences per branch, using ONLY the household income, income YoY, population YoY, and home value YoY figures given. Frame through Verlocity's AudienceFinder segments (High-Quality Local Prospects from income/geo, Regression-Scored Lookalikes, Competitive Conquesting for switchers, Warm Retargeting) where the demographic signal supports it. Never invent a named persona (e.g. 'Sarah, 34') -- Verlocity's demographic persona layer is in development, not live. Key must exactly match the branch name+city+state given."}
}"""

    if deep_dive_ctx:
        ctx += deep_dive_ctx

    print("  Generating AI narratives (full-network context)...")
    client = anthropic.Anthropic(api_key=ANTH_KEY)
    try:
        msg = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=4000,
            system=system,
            messages=[{"role": "user", "content": ctx}],
        )
        raw = msg.content[0].text.strip().replace("```json", "").replace("```", "").strip()
        narr = json.loads(raw)
        print("  ✓ Narratives generated")
        return narr
    except Exception as e:
        print(f"  ⚠ Narrative generation failed ({e}) — using placeholders")
        return _placeholder_narratives(dives)


def _placeholder_narratives(dives=None):
    base = {k: "" for k in ["exec_summary", "network_narrative", "competitive_narrative",
                             "financial_narrative", "capture_strategy_narrative", "next_step"]}
    base["branch_audiences"] = {}
    if dives:
        for e in dives:
            b = e["branch"]
            key = f"{b.get('namebr')} ({b.get('citybr')}, {b.get('stalpbr')})"
            base["branch_audiences"][key] = ""
    return base


# ═══════════════════════════════════════════════════════════════
# DOCX BUILD
# ═══════════════════════════════════════════════════════════════

def _set_cell_shading(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _heading(doc, text, size=16, color=NAVY, space_before=18, space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.name = FONT_HEAD  # Inter Bold per Verlocity_Brand_Guidelines_R2.pdf
    return p


def _body(doc, text, size=10.5, color=RGBColor(0x33, 0x33, 0x33)):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.name = FONT_HEAD  # Inter (Light per brand guide) — same family, body weight
    return p


def build_assessment_doc(bank_name, summary, fin, targets, narr, branches, branches_geo=None,
                          branch_strategy=None, dives=None, deep_mode=None, tmpdir="."):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.59)   # US Letter
    section.page_height = Cm(27.94)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)

    # ── Cover ──
    p0 = doc.add_paragraph()
    p0.paragraph_format.space_before = Pt(50)
    logo_path = str(Path(__file__).parent / "verlocity_logo.jpg") if (Path(__file__).parent / "verlocity_logo.jpg").exists() else None
    if logo_path:
        run0 = p0.add_run()
        run0.add_picture(logo_path, width=Inches(2.0))  # brand guide: min 1.0" width
    else:
        run0 = p0.add_run("VERLOCITY")
        run0.bold = True
        run0.font.size = Pt(14)
        run0.font.color.rgb = TEAL
        run0.font.name = FONT_HEAD

    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(18)
    run2 = p2.add_run(bank_name)
    run2.bold = True
    run2.font.size = Pt(28)
    run2.font.color.rgb = NAVY
    run2.font.name = FONT_HEAD

    p3 = doc.add_paragraph()
    run3 = p3.add_run("BMAP Market Assessment")
    run3.font.size = Pt(14)
    run3.font.color.rgb = GRAY3
    run3.font.name = FONT_HEAD

    p4 = doc.add_paragraph()
    p4.paragraph_format.space_after = Pt(30)
    run4 = p4.add_run(datetime.now().strftime("%B %Y"))
    run4.font.size = Pt(10)
    run4.font.color.rgb = GRAY3
    run4.font.name = FONT_HEAD

    # Cover hero visual — zone distribution, sets the McKinsey "big number up front" tone
    zone_chart_path = f"{tmpdir}/_chart_zones.png"
    chart_zone_distribution(summary["zones"], zone_chart_path)
    doc.add_picture(zone_chart_path, width=Inches(6.3))

    doc.add_page_break()

    # ── Executive Summary ──
    _heading(doc, "Executive Summary", space_before=0)

    # Pull-quote callout — the single most important number, McKinsey-style
    top_branch = summary["top5"][0] if summary["top5"] else None
    if top_branch:
        callout = doc.add_table(rows=1, cols=1)
        cell = callout.rows[0].cells[0]
        _set_cell_shading(cell, "083D5F")
        cell.paragraphs[0].text = ""
        r1 = cell.paragraphs[0].add_run(
            f"{top_branch['namebr']}: {_sf(top_branch['opportunity_score']):.0f}/100 opportunity score, "
            f"{_sf(top_branch['yoy_deposits'])*100:+.1f}% deposit growth"
        )
        r1.font.size = Pt(13)
        r1.font.bold = True
        r1.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        cell.paragraphs[0].paragraph_format.space_before = Pt(10)
        cell.paragraphs[0].paragraph_format.space_after = Pt(10)
        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    _body(doc, narr.get("exec_summary") or
          f"{bank_name} operates {summary['branch_count']} branches with ${summary['total_deposits_B']:.1f}B "
          f"in total deposits. Network average opportunity score: {summary['avg_score']:.0f}/100.")

    # ── Geographic Distribution (map) ──
    if branches_geo:
        map_path = f"{tmpdir}/_chart_map.png"
        if chart_branch_map(branches_geo, map_path):
            _heading(doc, "Geographic Distribution")
            doc.add_picture(map_path, width=Inches(6.3))

    # ── Network Opportunity Overview ──
    _heading(doc, "Network Opportunity Overview")
    _body(doc, narr.get("network_narrative") or "")

    if summary["top5"] and summary["bottom3"]:
        top_bottom_path = f"{tmpdir}/_chart_topbottom.png"
        chart_top_bottom_branches(summary["top5"], summary["bottom3"], top_bottom_path)
        p_lbl = doc.add_paragraph()
        p_lbl.paragraph_format.space_before = Pt(10)
        r = p_lbl.add_run("Highest- and Lowest-Opportunity Branches")
        r.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = NAVY
        doc.add_picture(top_bottom_path, width=Inches(6.3))

    # ── Competitive Overview ──
    _heading(doc, "Competitive Overview")
    _body(doc, narr.get("competitive_narrative") or "")
    if targets:
        ct = doc.add_table(rows=1, cols=4)
        ct.style = "Light Grid Accent 1"
        hdr = ct.rows[0].cells
        for i, h in enumerate(["Target", "Branches Exposed", "Vulnerability", "Deposit YoY"]):
            hdr[i].text = h
        for t in targets:
            row = ct.add_row().cells
            row[0].text = str(t.get("target_institution", "—"))
            row[1].text = str(t.get("branches_in_radius", "—"))
            row[2].text = _vuln_tier(t.get("avg_vuln_score"))
            row[3].text = f"{_sf(t.get('avg_yoy_pct')):+.1f}%"

    # ── Deposit Capture Strategy (adaptive-radius, branch + bank-wide) ──
    if branch_strategy:
        _heading(doc, "Deposit Capture Strategy — By Branch and Network-Wide")
        _body(doc, narr.get("capture_strategy_narrative") or
              "Competitive radius is scaled per branch by local density and deposit size, rather than "
              "a single fixed distance — dense, high-value branches are assessed as tight as 0.5 miles; "
              "rural or low-deposit branches widen to 10 miles, since a thin market can otherwise miss "
              "the few competitors that do exist.")

        bs_summary = summarize_branch_strategy(branch_strategy)
        rt = doc.add_table(rows=1, cols=3)
        rt.style = "Light Grid Accent 1"
        hdr = rt.rows[0].cells
        for i, h in enumerate(["Market Tier", "Branches", "Deposits in Tier"]):
            hdr[i].text = h
        for tier in ["Dense/High-Value", "Dense", "Suburban", "Low-Density"]:
            v = bs_summary["tiers"].get(tier)
            if not v:
                continue
            row = rt.add_row().cells
            row[0].text = tier
            row[1].text = str(v["count"])
            row[2].text = f"${v['deposits']/1e6:.0f}M"
        doc.add_paragraph().paragraph_format.space_after = Pt(10)

        # Scenario-based $ opportunity — low/medium/aggressive annual capture rate
        # against the contestable deposit pool per tier. Industry-informed planning
        # assumption, not Mid Penn-specific history -- flagged as such in-doc.
        p_lbl2 = doc.add_paragraph()
        p_lbl2.paragraph_format.space_before = Pt(6)
        r2 = p_lbl2.add_run("Projected Annual Capture by Scenario")
        r2.bold = True
        r2.font.size = Pt(11)
        r2.font.color.rgb = NAVY
        r2.font.name = FONT_HEAD

        st = doc.add_table(rows=1, cols=4)
        st.style = "Light Grid Accent 1"
        hdr = st.rows[0].cells
        for i, h in enumerate(["Market Tier", "Low (1%)", "Medium (3%)", "Aggressive (7%)"]):
            hdr[i].text = h
        for tier in ["Dense/High-Value", "Dense", "Suburban", "Low-Density"]:
            v = bs_summary["tiers"].get(tier)
            if not v:
                continue
            dep = v["deposits"]
            row = st.add_row().cells
            row[0].text = tier
            row[1].text = f"${dep*0.01/1e6:.1f}M"
            row[2].text = f"${dep*0.03/1e6:.1f}M"
            row[3].text = f"${dep*0.07/1e6:.1f}M"
        p_note = doc.add_paragraph()
        p_note.paragraph_format.space_before = Pt(4)
        note_run = p_note.add_run(
            "Industry-informed planning assumption, not Mid Penn-specific history. Retail deposit "
            "switching is behaviorally sticky; 7-8% annual capture of a local contestable pool is close "
            "to a realistic ceiling absent a genuine market disruption. Replace with the bank's own "
            "historical account-opening and deposit-capture data once available, per the same "
            "calibration principle used in the Predictive ROI model."
        )
        note_run.italic = True
        note_run.font.size = Pt(8.5)
        note_run.font.color.rgb = GRAY3
        doc.add_paragraph().paragraph_format.space_after = Pt(10)

        # Named per-branch findings — the actual capture targets, not just tier counts
        named = [r for r in branch_strategy if r.get("top_competitor")]
        named = sorted(named, key=lambda r: -r["deposits"])[:12]
        if named:
            p_lbl = doc.add_paragraph()
            r = p_lbl.add_run("Named Nearest Competitor by Branch (adaptive radius)")
            r.bold = True
            r.font.size = Pt(11)
            r.font.color.rgb = NAVY
            nt = doc.add_table(rows=1, cols=5)
            nt.style = "Light Grid Accent 1"
            hdr = nt.rows[0].cells
            for i, h in enumerate(["Branch", "Tier", "Radius", "Largest Nearby Competitor", "Distance"]):
                hdr[i].text = h
            for r_ in named:
                tc = r_["top_competitor"]
                row = nt.add_row().cells
                row[0].text = f"{r_['namebr']} ({r_['citybr']}, {r_['stalpbr']})"
                row[1].text = r_["tier"]
                row[2].text = f"{r_['radius_mi']}mi"
                row[3].text = str(tc.get("bank_name", "—"))
                row[4].text = f"{_sf(tc.get('distance_miles')):.2f}mi"

    # ── Branch Assessment — full 16-play deep dive ──
    # <20 branches: every branch assessed. >=20: curated to the 15 with the
    # biggest actual $ capture opportunity. dives/deep_mode computed once
    # upstream (shared with get_narratives for the audience blurbs) rather
    # than recomputed here, so narrative keys and doc content stay in sync.
    if dives is None:
        dives, deep_mode = build_branch_deep_dives(branches, branch_strategy or [])

    branch_audiences = narr.get("branch_audiences") or {}

    if dives:
        section_title = ("Branch-by-Branch Assessment" if deep_mode
                          else f"Priority Branch Deep Dives — Top {len(dives)} by Capture Opportunity")
        _heading(doc, section_title)
        if deep_mode:
            _body(doc, f"This network has {len(branches)} branches — under the {DEEP_DIVE_THRESHOLD}-branch "
                       f"threshold for full individual coverage. Every branch below is assessed on zone, "
                       f"competitive quadrant, assigned play, radius methodology, named competitors, and "
                       f"audience signal.")
        else:
            _body(doc, f"This network has {len(branches)} branches — above the threshold for full "
                       f"individual coverage. The {len(dives)} branches below are ranked by actual "
                       f"dollar capture opportunity (named competitor's deposits within the branch's "
                       f"adaptive radius, size-filtered to 0.1x-5x the branch's own deposits per the "
                       f"BMAP vulnerability methodology), not opportunity score alone.")
        doc.add_page_break()

        for i, e in enumerate(dives):
            b = e["branch"]
            strat = e["strategy"]
            play = e["play"]
            branch_label = f"{b.get('namebr','—')} ({b.get('citybr','—')}, {b.get('stalpbr','—')})"
            q_full = b.get("matrix_quadrant") or "—"
            q_short = q_full.split(" - ")[0]

            # Branch name as its own heading, colored by zone for quick scanning
            zone = b.get("opportunity_zone", "Analyze")
            zone_rgb = rgb(ZONE_COLOR.get(zone, "185FA5"))
            _heading(doc, branch_label, size=15, color=zone_rgb, space_before=(0 if i == 0 else 4))

            # ── Overview stat block ──
            ov = doc.add_table(rows=2, cols=4)
            ov.style = "Light Grid Accent 1"
            ov_hdr = ov.rows[0].cells
            for j, h in enumerate(["Zone", "Quadrant", "Priority Tier", "Opportunity Score"]):
                ov_hdr[j].text = h
            ov_val = ov.rows[1].cells
            ov_val[0].text = zone
            ov_val[1].text = q_full
            ov_val[2].text = str(b.get("priority_tier") or "—")
            ov_val[3].text = f"{_sf(b.get('opportunity_score')):.0f}/100"
            doc.add_paragraph().paragraph_format.space_after = Pt(4)

            # ── Deposits & Radius ──
            _heading(doc, "Deposits & Radius Methodology", size=11, space_before=8, space_after=4)
            dep = _sf(b.get("latest_dep"))
            yoy = _sf(b.get("yoy_deposits")) * 100
            radius_mi = strat.get("radius_mi") if strat else None
            tier_label = strat.get("tier") if strat else None
            dr = doc.add_table(rows=2, cols=4)
            dr.style = "Light Grid Accent 1"
            dr_hdr = dr.rows[0].cells
            for j, h in enumerate(["Deposits", "YoY Growth", "Market Tier", "Radius Used"]):
                dr_hdr[j].text = h
            dr_val = dr.rows[1].cells
            dr_val[0].text = f"${dep/1e6:.1f}M"
            dr_val[1].text = f"{yoy:+.1f}%"
            dr_val[2].text = tier_label or "—"
            dr_val[3].text = f"{radius_mi}mi" if radius_mi else "—"
            p_method = doc.add_paragraph()
            p_method.paragraph_format.space_before = Pt(4)
            method_run = p_method.add_run(
                "Radius scales to local density and deposit size — as tight as 0.5mi for "
                "dense, high-value markets, up to 10mi for rural or low-deposit branches — "
                "rather than one fixed distance for the whole network."
            )
            method_run.italic = True
            method_run.font.size = Pt(8.5)
            method_run.font.color.rgb = GRAY3
            method_run.font.name = FONT_HEAD

            # ── Competitors (top 3, size-filtered) ──
            _heading(doc, "Named Competitors", size=11, space_before=10, space_after=4)
            top3 = (strat.get("top3_competitors") if strat else []) or []
            if top3:
                cot = doc.add_table(rows=1, cols=4)
                cot.style = "Light Grid Accent 1"
                cot_hdr = cot.rows[0].cells
                for j, h in enumerate(["Competitor", "City / State", "Distance", "Deposits"]):
                    cot_hdr[j].text = h
                for c in top3:
                    row = cot.add_row().cells
                    row[0].text = str(c.get("bank_name", "—"))
                    row[1].text = f"{c.get('city','—')}, {c.get('state','—')}"
                    row[2].text = f"{_sf(c.get('distance_miles')):.2f}mi"
                    row[3].text = f"${_sf(c.get('deposits'))/1e6:.1f}M"
            else:
                _body(doc, "No competitor within the adaptive radius meets the 0.1x-5x size filter — "
                           "this branch has natural geographic protection rather than a capture target.",
                      size=9.5)

            # ── Capture Scenario ──
            capture_pool = e["capture_pool"]
            if capture_pool:
                _heading(doc, "Projected Annual Capture", size=11, space_before=10, space_after=4)
                cst = doc.add_table(rows=2, cols=3)
                cst.style = "Light Grid Accent 1"
                cst_hdr = cst.rows[0].cells
                for j, h in enumerate(["Low (1%)", "Medium (3%)", "Aggressive (7%)"]):
                    cst_hdr[j].text = h
                cst_val = cst.rows[1].cells
                cst_val[0].text = f"${capture_pool*0.01/1e6:.2f}M"
                cst_val[1].text = f"${capture_pool*0.03/1e6:.2f}M"
                cst_val[2].text = f"${capture_pool*0.07/1e6:.2f}M"

            # ── Audience Signal (real Census/ZHVI, no fabricated personas) ──
            audience_text = branch_audiences.get(branch_label, "")
            if audience_text or b.get("household_income"):
                _heading(doc, "Audience Signal", size=11, space_before=10, space_after=4)
                if audience_text:
                    _body(doc, audience_text, size=9.5)
                else:
                    inc = _sf(b.get("household_income"))
                    inc_yoy = _sf(b.get("yoy_income_growth")) * 100
                    pop_yoy = _sf(b.get("yoy_pop_growth")) * 100
                    zhvi_yoy = _sf(b.get("zhvi_yoy_pct"))  # already a percentage, not a decimal
                    _body(doc, f"Household income ${inc:,.0f} ({inc_yoy:+.1f}% YoY). Population growth "
                               f"{pop_yoy:+.1f}% YoY. Home values {zhvi_yoy:+.1f}% YoY.", size=9.5)

            # ── Play ──
            if play:
                _heading(doc, "Assigned Play", size=11, color=NAVY, space_before=10, space_after=4)
                play_box = doc.add_table(rows=1, cols=1)
                cell = play_box.rows[0].cells[0]
                _set_cell_shading(cell, "083D5F")
                cell.paragraphs[0].text = ""
                pr = cell.paragraphs[0].add_run(play.upper())
                pr.font.size = Pt(13)
                pr.font.bold = True
                pr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                pr.font.name = FONT_HEAD
                cell.paragraphs[0].paragraph_format.space_before = Pt(8)
                cell.paragraphs[0].paragraph_format.space_after = Pt(8)

                posture = PLAY_ACQUISITION_POSTURE.get(play, "")
                brief = PLAY_MEDIA_BRIEF.get(play, "")
                if posture:
                    p1 = doc.add_paragraph()
                    p1.paragraph_format.space_before = Pt(6)
                    lbl1 = p1.add_run("Resource posture: ")
                    lbl1.bold = True
                    lbl1.font.size = Pt(9.5)
                    lbl1.font.color.rgb = NAVY
                    lbl1.font.name = FONT_HEAD
                    r1 = p1.add_run(posture)
                    r1.font.size = Pt(9.5)
                    r1.font.name = FONT_HEAD
                if brief:
                    p2 = doc.add_paragraph()
                    p2.paragraph_format.space_after = Pt(4)
                    lbl2 = p2.add_run("Media brief: ")
                    lbl2.bold = True
                    lbl2.font.size = Pt(9.5)
                    lbl2.font.color.rgb = NAVY
                    lbl2.font.name = FONT_HEAD
                    r2 = p2.add_run(brief)
                    r2.font.size = Pt(9.5)
                    r2.font.name = FONT_HEAD

            if i < len(dives) - 1:
                doc.add_page_break()

    # ── Financial Health Benchmarking ──
    _heading(doc, "Financial Health Benchmarking")
    _body(doc, narr.get("financial_narrative") or "")

    fin_chart_path = f"{tmpdir}/_chart_financial.png"
    chart_financial_benchmark(fin, fin_chart_path)
    doc.add_picture(fin_chart_path, width=Inches(6.3))

    ft = doc.add_table(rows=1, cols=3)
    ft.style = "Light Grid Accent 1"
    hdr = ft.rows[0].cells
    for i, h in enumerate(["Metric", "Value", "Benchmark"]):
        hdr[i].text = h
    metrics = [
        ("ROA", f"{_sf(fin.get('roa')):.2f}%", ">1.0%"),
        ("NIM", f"{_sf(fin.get('nim')):.2f}%", "2.5–3.5%"),
        ("Efficiency Ratio", f"{_sf(fin.get('efficiency_ratio')):.1f}%", "<60%"),
        ("Deposit YoY", f"{_sf(fin.get('dep_yoy_pct')):+.1f}%", ">2%"),
        ("Cost of Funds", f"{_sf(fin.get('cost_of_funds_pct')):.2f}%", "<2%"),
        ("Tier 1 Capital", f"{_sf(fin.get('tier1_capital_pct')):.1f}%", ">8%"),
        ("Net Income YoY", f"{_sf(fin.get('net_income_yoy_pct')):+.1f}%", ">0%"),
    ]
    for label, val, bench in metrics:
        row = ft.add_row().cells
        row[0].text = label
        row[1].text = val
        row[2].text = bench

    # ── Next Step Recommendation ──
    _heading(doc, "Recommendation")
    _body(doc, narr.get("next_step") or "")

    doc.add_page_break()

    # ── Full Branch Appendix ──
    _heading(doc, f"Appendix — Full Branch Scoring ({len(branches)} branches)", space_before=0)
    at = doc.add_table(rows=1, cols=6)
    at.style = "Light Grid Accent 1"
    hdr = at.rows[0].cells
    for i, h in enumerate(["Branch", "City / State", "Deposits", "YoY", "Score", "Zone"]):
        hdr[i].text = h
    for b in branches:
        row = at.add_row().cells
        row[0].text = str(b.get("namebr", "—"))
        row[1].text = f"{b.get('citybr','—')}, {b.get('stalpbr','—')}"
        row[2].text = f"${_sf(b.get('latest_dep'))/1e6:.1f}M"
        row[3].text = f"{_sf(b.get('yoy_deposits'))*100:+.1f}%"
        row[4].text = f"{_sf(b.get('opportunity_score')):.0f}"
        row[5].text = str(b.get("opportunity_zone", "—"))

    # ── Session placeholder (deliberately not narrative content) ──
    doc.add_page_break()
    _heading(doc, "Discussed Live", space_before=0)
    _body(doc, "This assessment includes a working session to walk through these findings and answer "
               "specific questions about your network. Session notes and next-step scope are recorded separately.")

    return doc


def save_doc(doc, bank_name, out_dir=OUT_DIR):
    safe = "".join(c if c.isalnum() or c in " _-" else "_" for c in bank_name).strip()
    date = datetime.now().strftime("%Y%m%d")
    fname = out_dir / f"BMAP_Assessment_{safe}_{date}.docx"
    doc.save(str(fname))
    return fname


# ═══════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════

def run(ik, name_hint=None):
    print(f"\n{'='*60}\n  BMAP Assessment — {name_hint or ik}\n{'='*60}")
    d = fetch_full_network_data(ik)
    if not d["branches"]:
        raise ValueError(
            f"No branch data found for inst_key='{ik}'. This bank is either "
            f"not ingested into branch_opportunity_base, or the inst_key is "
            f"wrong. Refusing to generate a document — an empty-data report "
            f"would look identical to a real one and could ship by mistake."
        )
    bank_name = name_hint or (d["branches"][0].get("namefull") if d["branches"] else None) or ik
    summary = summarize_network(d)
    dives, deep_mode = build_branch_deep_dives(d["branches"], d.get("branch_strategy") or [])
    narr = get_narratives(bank_name, summary, d["fin"], d["targets"], d.get("branch_strategy"), dives)
    import tempfile
    with tempfile.TemporaryDirectory() as tmpdir:
        doc = build_assessment_doc(bank_name, summary, d["fin"], d["targets"], narr,
                                    d["branches"], d.get("branches_geo"),
                                    d.get("branch_strategy"), dives, deep_mode, tmpdir=tmpdir)
        path = save_doc(doc, bank_name)
    print(f"\n  ✓ Saved: {path}\n")
    return path


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--inst_key", required=True)
    ap.add_argument("--name", default=None)
    args = ap.parse_args()
    run(args.inst_key, args.name)
