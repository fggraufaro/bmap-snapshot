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

# Zone colors — matches BMAP_Snapshot_Hancock_Whitney_Bank_20260813.pptx exactly
# (brand-guideline navy→teal→emerald→lemon progression, not a red/green traffic-light scheme)
ZONE_COLOR = {
    "Invest":  "083D5F", "Analyze": "02A7C2", "Defend": "66CC99", "Justify": "BFC815",
}
ZONE_LIGHT = {
    "Invest":  "E1E9EF", "Analyze": "E1F5F8", "Defend": "EAF7F0", "Justify": "F8FAE3",
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
    "raw_sod":                        "raw",
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

    print(f"  Resolving winsorized YoY values...")
    capped_yoy = resolve_capped_yoy(branches)
    if capped_yoy:
        print(f"  ✓ {len(capped_yoy)} capped branch(es) resolved against raw_sod")

    return {
        "inst_key": ik,
        "branches": branches,
        "fin": fin,
        "targets": tgt_arr,
        "branches_geo": branches_geo,
        "branch_strategy": branch_strategy,
        "capped_yoy": capped_yoy,
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


def resolve_capped_yoy(branches):
    """yoy_deposits is winsorized/capped at 1.0 (100%) upstream -- a real,
    known limitation (documented pending item: 'real_yoy_display'). Every
    branch showing exactly 1.0 collapses two different real situations into
    one misleading '+100.0%': genuine outsized growth, or a branch with no
    prior-year row at all (new/newly-ingested branch, where YoY is undefined,
    not 100%). This resolves both cases directly from raw.raw_sod so the
    doc never publishes a fabricated-looking exact 100%.
    Returns {uninumbr: display_string} for capped branches only."""
    capped = [b for b in branches if abs(_sf(b.get("yoy_deposits")) - 1.0) < 0.001]
    if not capped:
        return {}

    ids = ",".join(str(b["uninumbr"]) for b in capped)
    rows = supabase("raw_sod", f"UNINUMBR=in.({ids})&select=UNINUMBR,YEAR,DEPSUMBR")
    by_branch = {}
    for r in rows if isinstance(rows, list) else []:
        by_branch.setdefault(str(r["UNINUMBR"]), {})[str(r["YEAR"])] = _sf(r.get("DEPSUMBR"))

    years = sorted({y for v in by_branch.values() for y in v.keys()})
    if len(years) < 2:
        cur_yr, prior_yr = (years[-1], None) if years else (None, None)
    else:
        cur_yr, prior_yr = years[-1], years[-2]

    out = {}
    for b in capped:
        uid = str(b["uninumbr"])
        vals = by_branch.get(uid, {})
        cur = vals.get(cur_yr)
        prior = vals.get(prior_yr) if prior_yr else None
        if prior and prior > 0 and cur is not None:
            real_pct = (cur - prior) / prior * 100
            out[b["uninumbr"]] = f"{real_pct:+.1f}%*"
        else:
            out[b["uninumbr"]] = "New branch*"
    return out


def fmt_yoy(b, capped_map):
    """Single formatter for yoy_deposits everywhere it's displayed --
    routes through the capped-value resolution instead of ever printing
    a bare '+100.0%'."""
    override = capped_map.get(b.get("uninumbr"))
    if override:
        return override
    return f"{_sf(b.get('yoy_deposits'))*100:+.1f}%"


# ═══════════════════════════════════════════════════════════════
# VISUALS — matplotlib charts + geographic map, embedded as PNGs.
# McKinsey-style: clean, minimal chrome, brand colors, direct labels
# instead of legends where possible.
# ═══════════════════════════════════════════════════════════════
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib.patheffects as pe

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


_STATES_GEOJSON_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "us_states.geojson")
_STATES_GEOJSON_URL = "https://raw.githubusercontent.com/PublicaMundi/MappingAPI/master/data/geojson/us-states.json"


def _load_state_polygons():
    """Local-first US state boundaries (Polygon/MultiPolygon per state).
    Ships as a repo asset (us_states.geojson) so report generation never
    depends on network access at request time; falls back to a one-time
    fetch + cache if the asset is missing."""
    import json as _json
    if not os.path.exists(_STATES_GEOJSON_PATH):
        try:
            import urllib.request
            with urllib.request.urlopen(_STATES_GEOJSON_URL, timeout=8) as resp:
                data = resp.read()
            with open(_STATES_GEOJSON_PATH, "wb") as f:
                f.write(data)
        except Exception:
            return {}
    try:
        with open(_STATES_GEOJSON_PATH) as f:
            gj = _json.load(f)
    except Exception:
        return {}

    from shapely.geometry import shape
    return {feat["properties"]["name"]: shape(feat["geometry"]) for feat in gj.get("features", [])}


def _draw_polygon(ax, geom, **kwargs):
    from matplotlib.patches import Polygon as MplPolygon
    from shapely.geometry import Polygon, MultiPolygon
    polys = geom.geoms if isinstance(geom, MultiPolygon) else [geom]
    for poly in polys:
        ax.add_patch(MplPolygon(list(poly.exterior.coords), closed=True, **kwargs))
        for interior in poly.interiors:
            ax.add_patch(MplPolygon(list(interior.coords), closed=True,
                                     facecolor="#FAFAF8", edgecolor="none", zorder=kwargs.get("zorder", 1) + 0.1))


def chart_branch_map_osm(branches_geo, path):
    """PRIMARY map — real street-map tiles (OpenStreetMap), the recognizable
    'Google Maps view' people actually orient by, instead of a bare state
    outline. Built with the staticmap package (pure Python, no GDAL, unlike
    contextily) which fetches and stitches raster tiles.

    Uses staticmap's own Web Mercator projection math (_lon_to_x/_lat_to_y +
    the instance's _x_to_px/_y_to_px after render) to place matplotlib labels
    at pixel-correct positions on top of the fetched tile image -- overlaying
    plain lon/lat coordinates on a Mercator-projected image without this
    conversion silently misplaces markers, worse the further from the map's
    vertical center.

    This function cannot be tested from within the dev sandbox (tile.
    openstreetmap.org is outside the sandbox's network allowlist) -- Railway
    has open egress, but this must be visually verified there. That's exactly
    why this has a defensive fallback in chart_branch_map(): any failure here
    (timeout, tile server error, missing package) falls back to the
    already-verified state-outline map rather than failing generation."""
    from staticmap import StaticMap, CircleMarker
    from staticmap.staticmap import _lon_to_x, _lat_to_y
    import math

    if not branches_geo:
        return False

    W, H = 1600, 1200
    m = StaticMap(W, H, padding_x=60, padding_y=60,
                  tile_request_timeout=8, delay_between_retries=0)
    for b in branches_geo:
        zone = b.get("opportunity_zone")
        color = ZONE_HEX_MPL.get(zone, "#778899")
        r_px = max(4, min(22, 4 + _sf(b.get("latest_dep")) / 5e7))
        m.add_marker(CircleMarker((b["lon"], b["lat"]), color, r_px))

    img = m.render()  # network call — the thing that can fail/time out

    # Attribution baked into the image itself (OSM tile usage policy),
    # not left to doc-building code that could later drop a caption.
    from PIL import ImageDraw
    draw = ImageDraw.Draw(img)
    draw.rectangle([0, H - 22, 260, H], fill=(255, 255, 255, 200))
    draw.text((6, H - 18), "Map data (c) OpenStreetMap contributors", fill=(60, 60, 60))

    # City labels for the largest branch per city — same top-8-by-count
    # logic as the fallback map, positioned via staticmap's own projection
    # so they land correctly on the Mercator tile image.
    by_city = {}
    for b in branches_geo:
        city = b.get("citybr") or b.get("city")
        if city:
            by_city.setdefault(city, []).append(b)
    top_cities = sorted(by_city.items(), key=lambda kv: -len(kv[1]))[:8]

    fig, ax = plt.subplots(figsize=(W / 200, H / 200), dpi=200)
    ax.imshow(img)
    for city, pts in top_cities:
        anchor = max(pts, key=lambda p: _sf(p.get("latest_dep")))
        px = m._x_to_px(_lon_to_x(anchor["lon"], m.zoom))
        py = m._y_to_px(_lat_to_y(anchor["lat"], m.zoom))
        ax.annotate(city, (px, py), xytext=(0, 12), textcoords="offset points",
                    fontsize=7.5, color="#1A1A1A", ha="center", va="top", zorder=4,
                    bbox=dict(boxstyle="round,pad=0.2", facecolor="white",
                              edgecolor="none", alpha=0.8))

    # Legend (matplotlib proxy handles, matching the fallback map's style)
    from matplotlib.lines import Line2D
    handles = [Line2D([0], [0], marker="o", linestyle="", markersize=7,
                       markerfacecolor=ZONE_HEX_MPL[z], markeredgecolor="white", label=z)
               for z in ["Justify", "Defend", "Analyze", "Invest"]]
    ax.legend(handles=handles, frameon=True, framealpha=0.9, fontsize=9,
              loc="lower left", edgecolor="none")

    ax.set_xticks([])
    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_visible(False)
    fig.tight_layout(pad=0)
    fig.savefig(path, dpi=200)
    plt.close(fig)
    return True


def chart_branch_map(branches_geo, path):
    """Dispatcher: real map tiles when available, state-outline fallback
    otherwise. See chart_branch_map_osm's docstring for why this needs a
    fallback and can't be verified from the dev sandbox."""
    try:
        if chart_branch_map_osm(branches_geo, path):
            return True
    except Exception as e:
        print(f"  ⚠ OSM map tiles failed ({e}) — falling back to state-outline map")
    return chart_branch_map_states(branches_geo, path)


def chart_branch_map_states(branches_geo, path):
    """FALLBACK map — lat/lon scatter, sized by deposits, colored by zone,
    drawn over real US state outlines cropped to the branch footprint. State
    boundaries come from a bundled GeoJSON asset (no live tile service needed).
    Used only if chart_branch_map_osm (real map tiles) fails for any reason —
    missing package, tile server down, network timeout on Railway — so a
    report never fails to generate over a map image."""
    if not branches_geo:
        return False

    lons_all = [b["lon"] for b in branches_geo]
    lats_all = [b["lat"] for b in branches_geo]
    pad_lon = max((max(lons_all) - min(lons_all)) * 0.18, 0.6)
    pad_lat = max((max(lats_all) - min(lats_all)) * 0.18, 0.6)
    x0, x1 = min(lons_all) - pad_lon, max(lons_all) + pad_lon
    y0, y1 = min(lats_all) - pad_lat, max(lats_all) + pad_lat

    fig, ax = plt.subplots(figsize=(7.2, 5.4), dpi=200)

    state_polys = _load_state_polygons()
    from shapely.geometry import box as _box
    view_box = _box(x0, y0, x1, y1)
    for name, geom in state_polys.items():
        if not geom.intersects(view_box):
            continue
        _draw_polygon(ax, geom, facecolor="#F2F2EE", edgecolor="#C7C7BF", linewidth=0.8, zorder=1)

    for zone in ["Justify", "Defend", "Analyze", "Invest"]:  # draw Invest last (on top)
        pts = [b for b in branches_geo if b.get("opportunity_zone") == zone]
        if not pts:
            continue
        lons = [b["lon"] for b in pts]
        lats = [b["lat"] for b in pts]
        sizes = [max(_sf(b.get("latest_dep")) / 3e6, 18) for b in pts]
        ax.scatter(lons, lats, s=sizes, c=ZONE_HEX_MPL[zone], alpha=0.85,
                   edgecolors="white", linewidths=0.5, label=zone, zorder=3)

    # Label the largest state clusters directly on the map (avg position, branch count)
    by_state = {}
    for b in branches_geo:
        st = b.get("stalpbr") or b.get("state")
        if st:
            by_state.setdefault(st, []).append(b)
    top_states = sorted(by_state.items(), key=lambda kv: -len(kv[1]))[:4]
    for st, pts in top_states:
        if len(pts) < 2:
            continue
        clx = sum(p["lon"] for p in pts) / len(pts)
        cly = sum(p["lat"] for p in pts) / len(pts)
        max_r = max(max(_sf(p.get("latest_dep")) / 3e6, 18) for p in pts)
        offset_pts = 14 + (max_r ** 0.5)  # clear large bubbles near the centroid
        ax.annotate(f"{st} · {len(pts)}", (clx, cly), xytext=(0, offset_pts),
                    textcoords="offset points", fontsize=8.5, fontweight="bold",
                    color="#334155", ha="center", zorder=4,
                    bbox=dict(boxstyle="round,pad=0.25", facecolor="#FAFAF8",
                              edgecolor="none", alpha=0.85))

    # City labels for individual reference points — state-level clustering
    # alone leaves a tight-footprint network (e.g. 22 branches all in one
    # county) with no landmark to orient by. Label the largest branch per
    # distinct city, capped to the top cities by branch count so a dense
    # network doesn't get cluttered with every city name.
    by_city = {}
    for b in branches_geo:
        city = b.get("citybr") or b.get("city")
        if city:
            by_city.setdefault(city, []).append(b)
    top_cities = sorted(by_city.items(), key=lambda kv: -len(kv[1]))[:8]
    for city, pts in top_cities:
        anchor = max(pts, key=lambda p: _sf(p.get("latest_dep")))
        r = max(_sf(anchor.get("latest_dep")) / 3e6, 18)
        offset_pts = 9 + (r ** 0.5) * 0.7
        ax.annotate(city, (anchor["lon"], anchor["lat"]), xytext=(0, -offset_pts),
                    textcoords="offset points", fontsize=6.5, color="#5B6472",
                    ha="center", va="top", zorder=4,
                    bbox=dict(boxstyle="round,pad=0.15", facecolor="#FAFAF8",
                              edgecolor="none", alpha=0.75))

    ax.set_xlim(x0, x1)
    ax.set_ylim(y0, y1)
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


def chart_branch_radius_map(branch_lat, branch_lon, competitors, radius_mi, path):
    """Small per-branch map: the branch at center, named competitors plotted
    at their real positions (not estimated bearings — branches_within_radius_batch
    now returns lat/lon directly), with a circle showing the adaptive radius
    actually used for that branch. Works in local miles-based x/y (not raw
    lon/lat degrees) so the radius circle renders as an actual circle instead
    of an ellipse — 1 degree of longitude covers fewer real miles than 1
    degree of latitude except at the equator, so plotting raw degrees under
    an equal-aspect axis distorts the shape."""
    if branch_lat is None or branch_lon is None:
        return False

    import math
    lat_rad = math.radians(branch_lat)
    mi_per_deg_lat = 69.0
    mi_per_deg_lon = 69.0 * max(math.cos(lat_rad), 0.15)

    def to_local_mi(lat, lon):
        return (lon - branch_lon) * mi_per_deg_lon, (lat - branch_lat) * mi_per_deg_lat

    fig, ax = plt.subplots(figsize=(3.4, 3.4), dpi=200)

    theta = [i / 100 * 2 * math.pi for i in range(101)]
    circ_x = [radius_mi * math.cos(t) for t in theta]
    circ_y = [radius_mi * math.sin(t) for t in theta]
    ax.plot(circ_x, circ_y, color="#083D5F", linewidth=1.0, linestyle="--", alpha=0.5, zorder=2)

    for c in competitors:
        clat, clon = c.get("lat"), c.get("lon")
        if clat is None or clon is None:
            continue
        cx, cy = to_local_mi(clat, clon)
        r = max(_sf(c.get("deposits")) / 4e6, 40)
        ax.scatter([cx], [cy], s=r, c="#A32D2D", alpha=0.75,
                   edgecolors="white", linewidths=0.6, zorder=3)
        label = c.get("bank_name", "")[:18]
        ax.annotate(f"{label}\n{_sf(c.get('distance_miles')):.1f}mi", (cx, cy),
                    xytext=(0, -9), textcoords="offset points", fontsize=5.5,
                    color="#334155", ha="center", va="top", zorder=4,
                    bbox=dict(boxstyle="round,pad=0.12", facecolor="#FAFAF8",
                              edgecolor="none", alpha=0.8))

    # The client's own branch, at the local origin, drawn last so it's on top
    ax.scatter([0], [0], s=140, c="#083D5F", marker="*",
               edgecolors="white", linewidths=0.8, zorder=5)

    pad = radius_mi * 1.35
    ax.set_xlim(-pad, pad)
    ax.set_ylim(-pad, pad)
    ax.set_xticks([])
    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.set_facecolor("#FAFAF8")
    ax.set_aspect("equal", adjustable="box")
    fig.tight_layout(pad=0.2)
    fig.savefig(path, transparent=False, facecolor="#FAFAF8")
    plt.close(fig)
    return True


def fetch_branch_geo(ik):
    """Lat/lon for the branch map — joined from geo.branches_master_v2.
    branch_id in geo matches uninumbr in branch_opportunity_base.
    stalpbr is required here (not just cosmetic) — chart_branch_map's
    cluster labeling groups by it, and silently labels nothing if it's
    missing rather than erroring, which is how this went unnoticed."""
    rows = supabase(
        "branch_opportunity_base",
        f"inst_key=eq.{ik}&select=uninumbr,opportunity_zone,latest_dep,stalpbr,citybr",
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

    # Demographic/audience aggregates -- BMAP's exec summary has been reporting
    # only opportunity score + financial ratios, but the platform also scores
    # every branch on Census income/population and ZHVI home-value trend
    # (the same signal AudienceFinder segments key off of). Without an
    # aggregate here, that whole data dimension never reaches the exec
    # summary — only individual branch_audiences blurbs get it.
    inc_vals = [_sf(b.get("household_income")) for b in br if b.get("household_income") is not None]
    inc_yoy_vals = [_sf(b.get("yoy_income_growth")) for b in br if b.get("yoy_income_growth") is not None]
    pop_yoy_vals = [_sf(b.get("yoy_pop_growth")) for b in br if b.get("yoy_pop_growth") is not None]
    zhvi_vals = [_sf(b.get("zhvi_yoy_pct")) for b in br if b.get("zhvi_yoy_pct") is not None]
    avg_household_income = sum(inc_vals) / len(inc_vals) if inc_vals else 0
    avg_income_yoy = sum(inc_yoy_vals) / len(inc_yoy_vals) if inc_yoy_vals else 0
    avg_pop_yoy = sum(pop_yoy_vals) / len(pop_yoy_vals) if pop_yoy_vals else 0
    avg_zhvi_yoy = sum(zhvi_vals) / len(zhvi_vals) if zhvi_vals else 0
    # The single branch with the strongest combined demographic tailwind
    # (income growth + population growth + home-value growth) -- gives the
    # exec summary a concrete named example instead of only network averages.
    strongest_demo_branch = None
    if br:
        strongest_demo_branch = max(
            br, key=lambda b: (_sf(b.get("yoy_income_growth")) + _sf(b.get("yoy_pop_growth"))
                                + _sf(b.get("zhvi_yoy_pct")) / 100)
        )

    top5 = sorted(br, key=lambda b: -_sf(b.get("opportunity_score")))[:5]
    bottom3 = sorted(br, key=lambda b: _sf(b.get("opportunity_score")))[:3]

    # Largest-deposit branch, tracked independent of opportunity-score rank.
    # top5/bottom3 above are pure score rankings, so a branch that dominates
    # the network by deposit size can rank outside both lists and never reach
    # the AI narrative context at all — which is exactly how a flagship branch
    # losing deposits can get buried under a small branch's growth story.
    largest_branch = max(br, key=lambda b: _sf(b.get("latest_dep"))) if br else None
    flagship_risk = None
    if largest_branch and total_dep > 0 and n > 0:
        share = _sf(largest_branch.get("latest_dep")) / total_dep
        yoy = _sf(largest_branch.get("yoy_deposits"))
        zone = largest_branch.get("opportunity_zone")
        # Material relative to network size (>=3x the average branch's share),
        # not a fixed absolute cutoff — a fixed 10% misses real cases on larger
        # networks (e.g. 9.7% share on a 59-branch network is ~5.7x average,
        # genuinely dominant, but would fail a flat 10% test).
        avg_share = 1.0 / n
        if share >= 3 * avg_share and (yoy < 0 or zone == "Justify"):
            flagship_risk = {**largest_branch, "deposit_share_pct": share * 100}

    return {
        "branch_count": n,
        "zones": zones,
        "total_deposits_B": total_dep / 1e9,
        "avg_yoy_pct": avg_yoy * 100,
        "avg_score": avg_score,
        "top5": top5,
        "bottom3": bottom3,
        "largest_branch": largest_branch,
        "flagship_risk": flagship_risk,
        "avg_household_income": avg_household_income,
        "avg_income_yoy_pct": avg_income_yoy * 100,
        "avg_pop_yoy_pct": avg_pop_yoy * 100,
        "avg_zhvi_yoy_pct": avg_zhvi_yoy,
        "strongest_demo_branch": strongest_demo_branch,
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


def get_narratives(bank_name, summary, fin, targets, branch_strategy=None, dives=None, capped_yoy=None):
    branch_strategy = branch_strategy or []
    dives = dives or []
    capped_yoy = capped_yoy or {}
    if not ANTH_KEY or not anthropic:
        print("  ⚠ No ANTHROPIC_API_KEY — using placeholder narratives")
        return _placeholder_narratives(dives)

    zones = summary["zones"]
    top5_str = "; ".join(
        f"{b['namebr']} ({b['citybr']}, {b['stalpbr']}) — score {_sf(b['opportunity_score']):.0f}, "
        f"${_sf(b['latest_dep'])/1e6:.0f}M deposits, {fmt_yoy(b, capped_yoy)} YoY"
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
{"FLAGSHIP RISK — the network's single largest branch by deposits carries the story regardless of its opportunity-score rank: " + summary['flagship_risk']['namebr'] + " (" + summary['flagship_risk']['citybr'] + ", " + summary['flagship_risk']['stalpbr'] + ") holds " + f"{summary['flagship_risk']['deposit_share_pct']:.0f}%" + " of total network deposits ($" + f"{_sf(summary['flagship_risk']['latest_dep'])/1e6:.0f}M" + "), is down " + f"{_sf(summary['flagship_risk']['yoy_deposits'])*100:+.1f}%" + " YoY, and sits in the " + str(summary['flagship_risk']['opportunity_zone']) + " zone." if summary.get('flagship_risk') else ""}

Financial health: ROA {_sf(fin.get('roa')):.2f}% | NIM {_sf(fin.get('nim')):.2f}% | Efficiency {_sf(fin.get('efficiency_ratio')):.1f}%
Deposit YoY {_sf(fin.get('dep_yoy_pct')):+.1f}% | Cost of funds {_sf(fin.get('cost_of_funds_pct')):.2f}% | Tier 1 {_sf(fin.get('tier1_capital_pct')):.1f}%
Net income YoY {_sf(fin.get('net_income_yoy_pct')):+.1f}%

Network-wide demographic & audience signal (Census income/population + ZHVI home-value
trend -- the same underlying data AudienceFinder segments key off of):
Avg household income ${summary['avg_household_income']:,.0f} | Avg income YoY {summary['avg_income_yoy_pct']:+.1f}%
Avg population YoY {summary['avg_pop_yoy_pct']:+.1f}% | Avg home-value (ZHVI) YoY {summary['avg_zhvi_yoy_pct']:+.1f}%
{"Strongest demographic tailwind: " + summary['strongest_demo_branch']['namebr'] + " (" + summary['strongest_demo_branch']['citybr'] + ", " + summary['strongest_demo_branch']['stalpbr'] + ") — income YoY " + f"{_sf(summary['strongest_demo_branch'].get('yoy_income_growth'))*100:+.1f}%" + ", population YoY " + f"{_sf(summary['strongest_demo_branch'].get('yoy_pop_growth'))*100:+.1f}%" + ", home value YoY " + f"{_sf(summary['strongest_demo_branch'].get('zhvi_yoy_pct')):+.1f}%" + "." if summary.get('strongest_demo_branch') else ""}

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

    system = """You are writing the $10K Verlocity BMAP Assessment for a community bank CEO/Board audience.
Tone: confident, commercial, decisive — not hedged, not academic. State the position, don't survey options.
Every branch reference MUST include both branch name AND city — never one without the other.
Do not explain BMAP methodology, do not reference BMAP versions, do not ask follow-up questions,
do not introduce data beyond what's given below. No superlatives for their own sake — earn every claim with a number.
Return ONLY valid JSON, no markdown fences:
{
  "exec_headline": "3-4 sentences. State plainly whether this network's current footprint is positioned for GROWTH, DEFENSE, or OPTIMIZATION -- pick one framing and commit to it. Reference the overall opportunity score and the Invest/Analyze/Defend/Justify mix. Frame the core deposit-acquisition tension explicitly (e.g. concentrated upside vs. broad retention burden). If a FLAGSHIP RISK finding is present above, it must anchor this headline by name and city -- it drives the network total and outweighs a smaller branch's score even if that branch tops the ranking.",
  "strategic_positioning": "One short paragraph (3-4 sentences). Describe what kind of deposit competitor this bank is positioned to be over the next 12-24 months. Ground this in the demographic & audience signal given (name the strongest-tailwind branch/city or the network averages -- growth demographics = expansion case, decline = retention case) AND the competitive/rate exposure (name the top vulnerable target). Address how growth can be driven without relying on additional physical branches -- position digital-first, market-specific execution as the default lever, without naming specific channels, tactics, or products.",
  "priority_focus": [{"branch": "exact branch name", "city": "city", "state": "ST", "zone": "Invest/Analyze/Defend/Justify", "why_now": "one clause: momentum, competitive pressure, or market structure -- with a number", "role": "one short strategic-role phrase, e.g. 'deposit growth engine', 'selective digital capture', 'defend and retain balances'"}] , // 2-3 entries. If a FLAGSHIP RISK finding is present, it MUST be one of these entries (role should reflect its risk, e.g. 'stabilize and retain' or 'exit review'). Otherwise lead with the top opportunity-score branch.
  "next_12_months": ["exactly 3 strings — each a leadership-level decision about WHERE to allocate attention, capital, or effort. No tactics, channels, offers, pricing, or products. No methodology."],
  "network_narrative": "2-3 sentences on what the zone distribution reveals about the network's overall position. (Used later in the doc, not the exec summary above -- can restate the zone framing in different words.)",
  "competitive_narrative": "2-3 sentences naming the specific network-level target and why it is vulnerable.",
  "financial_narrative": "2-3 sentences on what the financial metrics mean together — not a list restated as prose.",
  "capture_strategy_narrative": "3-4 sentences on the branch-level adaptive-radius findings. Name at least one specific dense/high-value branch with its named largest nearby competitor and distance, and contrast the tactical approach that implies (rate/digital competition at close range) against what the low-density branches need instead (defense and wallet-share deepening, since there is often no competitor within the adaptive radius to capture from). This is the 'win deposits by branch AND as a full bank' section.",
  "next_step": "2-3 sentences. A specific, named recommendation tied to the top opportunity branches. (Used in the closing Recommendation section, not the exec summary above.)",
  "branch_audiences": {"Branch Name (City, ST)": "2-3 sentences per branch, using ONLY the household income, income YoY, population YoY, and home value YoY figures given. Frame through Verlocity's AudienceFinder segments (High-Quality Local Prospects from income/geo, Regression-Scored Lookalikes, Competitive Conquesting for switchers, Warm Retargeting) where the demographic signal supports it. Never invent a named persona (e.g. 'Sarah, 34') -- Verlocity's demographic persona layer is in development, not live. Key must exactly match the branch name+city+state given."}
}"""

    if deep_dive_ctx:
        ctx += deep_dive_ctx

    print("  Generating AI narratives (full-network context)...")
    client = anthropic.Anthropic(api_key=ANTH_KEY)
    try:
        msg = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=8000,
            thinking={"type": "disabled"},
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
    base = {k: "" for k in ["exec_headline", "strategic_positioning", "network_narrative",
                             "competitive_narrative", "financial_narrative",
                             "capture_strategy_narrative", "next_step"]}
    base["priority_focus"] = []
    base["next_12_months"] = []
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
                          branch_strategy=None, dives=None, deep_mode=None, tmpdir=".", capped_yoy=None):
    capped_yoy = capped_yoy or {}
    geo_by_uid = {g["uninumbr"]: g for g in (branches_geo or []) if g.get("uninumbr") is not None}
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
            f"{fmt_yoy(top_branch, capped_yoy)} deposit growth"
        )
        r1.font.size = Pt(13)
        r1.font.bold = True
        r1.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        cell.paragraphs[0].paragraph_format.space_before = Pt(10)
        cell.paragraphs[0].paragraph_format.space_after = Pt(10)
        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    fallback_headline = (
        f"{bank_name} operates {summary['branch_count']} branches with ${summary['total_deposits_B']:.1f}B "
        f"in total deposits, network average opportunity score {summary['avg_score']:.0f}/100 across "
        f"{summary['zones']['Invest']} Invest, {summary['zones']['Analyze']} Analyze, "
        f"{summary['zones']['Defend']} Defend, and {summary['zones']['Justify']} Justify branches."
    )
    _body(doc, narr.get("exec_headline") or fallback_headline)

    if narr.get("strategic_positioning"):
        p_pos = doc.add_paragraph()
        p_pos.paragraph_format.space_before = Pt(8)
        r_pos = p_pos.add_run(narr["strategic_positioning"])
        r_pos.font.size = Pt(10.5)
        r_pos.font.name = FONT_HEAD
        r_pos.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    # Priority Focus — 2-3 named branches, structured (not a table, per the
    # source Executive Headline spec this was adapted from: name+city, zone,
    # why now, strategic role).
    priority_focus = narr.get("priority_focus") or []
    if not priority_focus:
        fr_local = summary.get("flagship_risk")
        for cand in ([fr_local] if fr_local else []) + summary["top5"][:3]:
            if not cand:
                continue
            already = any(pf.get("branch") == cand.get("namebr") for pf in priority_focus)
            if already or len(priority_focus) >= 3:
                continue
            is_risk = fr_local and cand.get("namebr") == fr_local.get("namebr")
            priority_focus.append({
                "branch": cand.get("namebr"), "city": cand.get("citybr"), "state": cand.get("stalpbr"),
                "zone": cand.get("opportunity_zone"),
                "why_now": f"{fmt_yoy(cand, capped_yoy)} YoY, {_sf(cand.get('opportunity_score')):.0f}/100 score.",
                "role": "stabilize and retain" if is_risk else "deposit growth engine",
            })
    if priority_focus:
        _heading(doc, "Priority Focus", size=12, space_before=12, space_after=4)
        for item in priority_focus:
            zone = item.get("zone", "")
            zone_rgb = rgb(ZONE_COLOR.get(zone, "083D5F"))
            p_pf = doc.add_paragraph()
            p_pf.paragraph_format.space_before = Pt(6)
            r_name = p_pf.add_run(f"{item.get('branch','—')} ({item.get('city','—')}, {item.get('state','—')}) ")
            r_name.font.bold = True
            r_name.font.size = Pt(10.5)
            r_name.font.name = FONT_HEAD
            r_name.font.color.rgb = NAVY
            r_zone = p_pf.add_run(f"· {zone}")
            r_zone.font.bold = True
            r_zone.font.size = Pt(9.5)
            r_zone.font.name = FONT_HEAD
            r_zone.font.color.rgb = zone_rgb
            p_pf2 = doc.add_paragraph()
            p_pf2.paragraph_format.space_after = Pt(2)
            r_role = p_pf2.add_run(f"{item.get('role','—')} — ")
            r_role.italic = True
            r_role.font.size = Pt(9.5)
            r_role.font.name = FONT_HEAD
            r_role.font.color.rgb = GRAY3
            r_why = p_pf2.add_run(item.get("why_now", ""))
            r_why.font.size = Pt(9.5)
            r_why.font.name = FONT_HEAD
            r_why.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    # What This Means for the Next 12 Months — exactly 3 leadership-level
    # decisions, no tactics/channels/pricing per spec.
    next_12 = narr.get("next_12_months") or []
    if not next_12:
        next_12 = [
            f"Allocate capital toward the {summary['zones']['Invest']} Invest-zone branches before "
            f"broad-based network spend.",
            f"Apply retention discipline across the {summary['zones']['Defend'] + summary['zones']['Justify']} "
            f"Defend/Justify branches rather than treating them as growth targets.",
            "Fund digital-first, market-specific execution as the primary lever for deposit growth "
            "beyond the existing physical footprint.",
        ]
    if next_12:
        _heading(doc, "What This Means for the Next 12 Months", size=12, space_before=12, space_after=4)
        for bullet in next_12:
            p_b = doc.add_paragraph(style="List Bullet")
            r_b = p_b.add_run(bullet)
            r_b.font.size = Pt(10)
            r_b.font.name = FONT_HEAD
            r_b.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    # Flagship-risk alert — guaranteed regardless of AI narrative compliance.
    # The pull-quote above is the top opportunity-score branch, which can be
    # a small branch; if the network's largest branch by deposits is
    # materially at risk, that fact belongs on page one too, not just in the
    # 59-row appendix table. (Priority Focus above is instructed to include
    # it by name when present, but this stays as a guaranteed backstop.)
    fr = summary.get("flagship_risk")
    if fr:
        alert = doc.add_table(rows=1, cols=1)
        cell = alert.rows[0].cells[0]
        _set_cell_shading(cell, "BFC815")  # brand Justify color
        cell.paragraphs[0].text = ""
        r_alert = cell.paragraphs[0].add_run(
            f"FLAGSHIP RISK — {fr['namebr']} ({fr['citybr']}, {fr['stalpbr']}) holds "
            f"{fr['deposit_share_pct']:.0f}% of total network deposits (${_sf(fr['latest_dep'])/1e6:.0f}M) "
            f"and is at {fmt_yoy(fr, capped_yoy)} YoY, {fr['opportunity_zone']} zone."
        )
        r_alert.font.size = Pt(12)
        r_alert.font.bold = True
        r_alert.font.color.rgb = NAVY
        cell.paragraphs[0].paragraph_format.space_before = Pt(8)
        cell.paragraphs[0].paragraph_format.space_after = Pt(8)
        doc.add_paragraph().paragraph_format.space_after = Pt(6)

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
            radius_mi = strat.get("radius_mi") if strat else None
            tier_label = strat.get("tier") if strat else None
            dr = doc.add_table(rows=2, cols=4)
            dr.style = "Light Grid Accent 1"
            dr_hdr = dr.rows[0].cells
            for j, h in enumerate(["Deposits", "YoY Growth", "Market Tier", "Radius Used"]):
                dr_hdr[j].text = h
            dr_val = dr.rows[1].cells
            dr_val[0].text = f"${dep/1e6:.1f}M"
            dr_val[1].text = fmt_yoy(b, capped_yoy)
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

            if top3:
                own = geo_by_uid.get(b.get("uninumbr"))
                if own and own.get("lat") is not None and own.get("lon") is not None:
                    radius_img = os.path.join(tmpdir, f"radius_{b.get('uninumbr')}.png")
                    ok = chart_branch_radius_map(
                        own["lat"], own["lon"], top3, radius_mi or 3.0, radius_img
                    )
                    if ok:
                        p_img = doc.add_paragraph()
                        p_img.paragraph_format.space_before = Pt(4)
                        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_img.add_run().add_picture(radius_img, width=Inches(2.4))

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
    fallback_next_step = None
    if not narr.get("next_step"):
        top = summary["top5"][0] if summary["top5"] else None
        fr = summary.get("flagship_risk")
        parts = []
        if fr:
            parts.append(f"Address {fr['namebr']}'s {fmt_yoy(fr, capped_yoy)} YoY position first — "
                          f"it holds {fr['deposit_share_pct']:.0f}% of total network deposits and outweighs "
                          f"any single opportunity-zone branch in dollar impact.")
        if top:
            parts.append(f"Prioritize capital toward {top['namebr']} ({top['citybr']}, {top['stalpbr']}), "
                          f"the network's top-scored branch at {_sf(top['opportunity_score']):.0f}/100.")
        parts.append(f"With {summary['zones']['Justify']} branches in Justify and "
                      f"{summary['zones']['Invest']} in Invest, the near-term agenda is reallocating "
                      f"capacity from the former to the latter.")
        fallback_next_step = " ".join(parts)
    _body(doc, narr.get("next_step") or fallback_next_step)

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
        row[3].text = fmt_yoy(b, capped_yoy)
        row[4].text = f"{_sf(b.get('opportunity_score')):.0f}"
        row[5].text = str(b.get("opportunity_zone", "—"))

    if capped_yoy:
        fn = doc.add_paragraph()
        fn.paragraph_format.space_before = Pt(6)
        fn_run = fn.add_run(
            "* YoY growth for this branch hit the standard calculation's cap and was "
            "resolved directly against source deposit data: shown as the real computed "
            "growth rate where a prior-year figure exists, or as \"New branch\" where none does."
        )
        fn_run.italic = True
        fn_run.font.size = Pt(8.5)
        fn_run.font.color.rgb = GRAY3
        fn_run.font.name = FONT_HEAD

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
    narr = get_narratives(bank_name, summary, d["fin"], d["targets"], d.get("branch_strategy"), dives,
                           d.get("capped_yoy"))
    import tempfile
    with tempfile.TemporaryDirectory() as tmpdir:
        doc = build_assessment_doc(bank_name, summary, d["fin"], d["targets"], narr,
                                    d["branches"], d.get("branches_geo"),
                                    d.get("branch_strategy"), dives, deep_mode, tmpdir=tmpdir,
                                    capped_yoy=d.get("capped_yoy"))
        path = save_doc(doc, bank_name)
    print(f"\n  ✓ Saved: {path}\n")
    return path


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--inst_key", required=True)
    ap.add_argument("--name", default=None)
    args = ap.parse_args()
    run(args.inst_key, args.name)
