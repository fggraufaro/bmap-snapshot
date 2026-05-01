"""
BMAP Snapshot Deck Builder
==========================
Generates a branded 6-slide PowerPoint deck for any bank in BMAP.
Uses python-pptx for native vector charts (crisp at any zoom).
Opens directly in PowerPoint — no Google Slides workaround needed.

Setup (one time):
    pip install python-pptx anthropic requests

Usage — single bank:
    python bmap_snapshot.py --inst_key bank_123

Usage — batch from CSV:
    python bmap_snapshot.py --csv banks.csv

CSV format (one bank per line):
    inst_key,optional_name
    bank_123,Country Bank for Savings
    bank_456

Output: BMAP_Snapshot_<BankName>_<date>.pptx  (or a folder for batch)
"""

import argparse
import csv
import io
import json
import os
import sys
import time
import urllib.request
from datetime import datetime
from pathlib import Path

# ── Third-party ────────────────────────────────────────────────
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
    from pptx.chart.data import ChartData
    from pptx.oxml.ns import qn
    import pptx.oxml as pxml
    from lxml import etree
except ImportError:
    sys.exit("Run: pip install python-pptx")

try:
    import anthropic
except ImportError:
    sys.exit("Run: pip install anthropic")

try:
    import requests
except ImportError:
    sys.exit("Run: pip install requests")

# ═══════════════════════════════════════════════════════════════
# CONFIG — edit these
# ═══════════════════════════════════════════════════════════════
SUPA_URL  = "https://tuiiywphoynbmkxpoyps.supabase.co"
SUPA_KEY  = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InR1aWl5d3Bob3luYm1reHBveXBzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NTc0MDg0NTMsImV4cCI6MjA3Mjk4NDQ1M30.8-JAz4WQRE3Fi6uH7xiYNTns92g-nV1A9pbUvSK549M"
ANTH_KEY  = os.environ.get("ANTHROPIC_API_KEY", "")  # set env var or paste here
LOGO_URL  = "https://fggraufaro.github.io/bmap-tools/Verlocity-Logo.png"
OUT_DIR   = Path(".")

# ═══════════════════════════════════════════════════════════════
# BRAND
# ═══════════════════════════════════════════════════════════════
def rgb(h): return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

NAVY    = rgb("1A2332")
TEAL    = rgb("1D9E75")
AMBER   = rgb("F5A623")
WHITE   = rgb("FFFFFF")
GRAY1   = rgb("F5F5F2")
GRAY2   = rgb("E8E8E5")
GRAY3   = rgb("778899")
INVEST  = rgb("27500A"); INVEST_L = rgb("EAF3DE")
ANALYZE = rgb("185FA5"); ANALYZE_L = rgb("E6F1FB")
DEFEND  = rgb("854F0B"); DEFEND_L  = rgb("FFF3E0")
JUSTIFY = rgb("A32D2D"); JUSTIFY_L = rgb("FCEBEB")

ZONE_C = {"Invest": INVEST, "Analyze": ANALYZE, "Defend": DEFEND, "Justify": JUSTIFY}
ZONE_L = {"Invest": INVEST_L, "Analyze": ANALYZE_L, "Defend": DEFEND_L, "Justify": JUSTIFY_L}

# Slide dimensions: 10" × 5.625" (widescreen)
W = Inches(10)
H = Inches(5.625)

# ═══════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════
def supabase(table, params):
    url = f"{SUPA_URL}/rest/v1/{table}?{params}"
    r = requests.get(url, headers={"apikey": SUPA_KEY, "Authorization": f"Bearer {SUPA_KEY}"}, timeout=30)
    r.raise_for_status()
    return r.json()

def add_rect(slide, x, y, w, h, fill_color, line_color=None, line_width=Pt(0)):
    shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid(); shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, text, x, y, w, h, size=11, bold=False, color=NAVY,
             align=PP_ALIGN.LEFT, italic=False, font="Calibri", valign="top"):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text)
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return tb

def add_chrome(slide, page_num, label, logo_bytes):
    """Verlocity sidebar + logo + section label + page number"""
    add_rect(slide, 0, 0, 0.28, 5.625, NAVY)
    add_rect(slide, 0.28, 0, 0.08, 5.625, TEAL)
    if logo_bytes:
        slide.shapes.add_picture(io.BytesIO(logo_bytes), Inches(0.42), Inches(5.04), Inches(1.55), Inches(0.36))
    if label:
        pill = add_rect(slide, 8.16, 0.14, 1.76, 0.30, TEAL)
        add_text(slide, label, 8.16, 0.14, 1.76, 0.30, size=7.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(slide, str(page_num), 9.50, 5.28, 0.38, 0.20, size=9, color=GRAY3, align=PP_ALIGN.RIGHT)

def add_narrative(slide, n, y0):
    """Left narrative column: headline / rule / spoken / bullets / close bar"""
    add_text(slide, n.get("headline",""), 0.45, y0, 5.6, 0.78,
             size=24, bold=True, color=NAVY, valign="bottom")
    add_rect(slide, 0.45, y0+0.84, 5.6, 0.04, TEAL)
    add_text(slide, n.get("spoken",""), 0.45, y0+0.96, 5.6, 0.60,
             size=9.5, italic=True, color=GRAY3)
    bullets = n.get("bullets", [])
    if bullets:
        tb = slide.shapes.add_textbox(Inches(0.45), Inches(y0+1.62), Inches(5.6), Inches(1.28))
        tf = tb.text_frame; tf.word_wrap = True
        for i, b in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = f"• {b}"
            p.font.size = Pt(9.5); p.font.color.rgb = NAVY; p.font.name = "Calibri"
            p.space_after = Pt(5)
    add_rect(slide, 0.45, y0+2.98, 5.6, 0.46, NAVY)
    add_text(slide, n.get("close",""), 0.56, y0+2.98, 5.4, 0.46,
             size=9.5, bold=True, color=WHITE)

def fetch_logo():
    try:
        with urllib.request.urlopen(LOGO_URL, timeout=10) as r:
            return r.read()
    except Exception as e:
        print(f"  ⚠  Logo fetch failed ({e}) — slides will have no logo")
        return None

# ═══════════════════════════════════════════════════════════════
# DATA FETCH
# ═══════════════════════════════════════════════════════════════
def fetch_bank_data(ik):
    print(f"  Fetching branch data...")
    rows = supabase("branch_opportunity_base",
        f"inst_key=eq.{ik}&select=namefull,latest_dep,yoy_deposits,avg_comp_yoy,"
        "opportunity_score,opportunity_zone,market_growth_score,inv_density_norm_winsor")

    print(f"  Fetching branch details...")
    br = supabase("branch_opportunity_base",
        f"inst_key=eq.{ik}&select=uninumbr,namebr,citybr,stalpbr,latest_dep,"
        "yoy_deposits,opportunity_score,opportunity_zone,matrix_quadrant,priority_tier"
        "&order=opportunity_score.desc&limit=50")

    print(f"  Fetching network target...")
    tgt_arr = supabase("vw_network_top_targets",
        f"my_inst_key=eq.{ik}&select=target_institution,branches_in_radius,"
        "avg_vuln_score,avg_yoy_pct,target_roa,target_efficiency_ratio,dominant_zone"
        "&order=network_rank.asc&limit=1")

    print(f"  Fetching financials...")
    fin_arr = supabase("bank_financial_snapshot_latest",
        f"inst_key=eq.{ik}&select=*&limit=1")

    return {
        "ik":       ik,
        "bankName": rows[0].get("namefull", ik) if rows else ik,
        "rows":     rows or [],
        "br":       br   or [],
        "tgt":      tgt_arr[0] if tgt_arr else None,
        "fin":      fin_arr[0] if fin_arr else {},
    }

# ═══════════════════════════════════════════════════════════════
# AI NARRATIVES
# ═══════════════════════════════════════════════════════════════
def get_narratives(data):
    if not ANTH_KEY:
        print("  ⚠  No ANTHROPIC_API_KEY — using placeholder narratives")
        empty = {"headline":"","spoken":"","bullets":[],"close":""}
        return {k: empty for k in ["network","priority","financial","nextsteps"]}

    rows = data["rows"]; fin = data["fin"]; tgt = data["tgt"]
    br   = data["br"];   bankName = data["bankName"]

    def sf(v, default=0):
        try:
            return float(v) if v is not None else default
        except (TypeError, ValueError):
            return default

    tot      = sum(sf(r.get("latest_dep")) for r in rows)
    avg      = lambda v: sum(sf(r.get(v)) for r in rows) / max(len(rows),1)
    invest   = sum(1 for r in rows if r.get("opportunity_zone")=="Invest")
    analyze  = sum(1 for r in rows if r.get("opportunity_zone")=="Analyze")
    defend   = sum(1 for r in rows if r.get("opportunity_zone")=="Defend")
    justify  = sum(1 for r in rows if r.get("opportunity_zone")=="Justify")
    bankYoY  = avg("yoy_deposits")*100
    compYoY  = avg("avg_comp_yoy")*100
    gap      = bankYoY - compYoY
    avgScore = avg("opportunity_score")

    top3 = sorted(br, key=lambda b: sf(b.get("opportunity_score")), reverse=True)[:3]
    top3_str = "; ".join(
        f"{b['namebr'].split('--')[-1].strip()} "
        f"(${sf(b.get('latest_dep'))/1e6:.0f}M, "
        f"{sf(b.get('yoy_deposits'))*100:.1f}% YoY, "
        f"score {sf(b.get('opportunity_score')):.0f})"
        for b in top3
    )

    ctx = f"""Bank: {bankName} | {len(rows)} branches | ${tot/1e9:.2f}B deposits
Deposit YoY: +{bankYoY:.1f}% | Peer avg: +{compYoY:.1f}% | Gap: {gap:+.1f}pp
Avg opp score: {avgScore:.1f}/100 | Zones: Invest {invest} | Analyze {analyze} | Defend {defend} | Justify {justify}
ROA: {fin.get('roa','—')}% | NIM: {fin.get('nim','—')}% | Efficiency: {fin.get('efficiency_ratio','—')}%
Net income YoY: {fin.get('net_income_yoy_pct','—')}% | Tier 1: {fin.get('tier1_capital_pct','—')}%
Top competitor: {tgt['target_institution']+' — '+str(tgt['branches_in_radius'])+' overlap branches' if tgt else 'N/A'}
Top 3 branches: {top3_str}"""

    system = """You are BMAP Executive Strategist at Verlocity. Write boardroom-quality slide narratives grounded in exact numbers.
Return ONLY valid JSON — no markdown, no explanation:
{"slides":[
  {"id":"network","headline":"strong claim max 9 words","spoken":"2 sentences Tom says walking in. Specific numbers.","bullets":["data point with number","competitive insight","risk or opportunity"],"close":"one punchy forward action with metric"},
  {"id":"priority","headline":"...","spoken":"...","bullets":[...],"close":"..."},
  {"id":"financial","headline":"...","spoken":"...","bullets":[...],"close":"..."},
  {"id":"nextsteps","headline":"...","spoken":"...","bullets":[...],"close":"..."}
]}"""

    print(f"  Generating AI narratives...")
    client = anthropic.Anthropic(api_key=ANTH_KEY)
    msg = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=2400,
        system=system,
        messages=[{"role":"user","content":f"Generate 4-slide narratives for:\n\n{ctx}"}]
    )
    raw = msg.content[0].text.strip().replace("```json","").replace("```","").strip()
    narr = {}
    try:
        for s in json.loads(raw)["slides"]:
            narr[s["id"]] = s
    except Exception as e:
        print(f"  ⚠  JSON parse failed ({e}) — using empty narratives")
    N = lambda k: narr.get(k, {"headline":"","spoken":"","bullets":[],"close":""})
    return {k: N(k) for k in ["network","priority","financial","nextsteps"]}

# ═══════════════════════════════════════════════════════════════
# SLIDE BUILDERS
# ═══════════════════════════════════════════════════════════════

def build_cover(prs, d, logo_bytes):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    slide.shapes.title and slide.shapes.title.element.getparent().remove(slide.shapes.title.element)

    add_rect(slide, 0, 0, 0.28, 5.625, NAVY)
    add_rect(slide, 0.28, 0, 0.08, 5.625, TEAL)

    # Logo top
    if logo_bytes:
        slide.shapes.add_picture(io.BytesIO(logo_bytes), Inches(1.2), Inches(0.28), Inches(2.45), Inches(0.54))

    # Teal rule
    add_rect(slide, 1.2, 0.96, 8.6, 0.05, TEAL)

    # Bank name
    add_text(slide, d["bankName"], 1.2, 1.1, 8.5, 0.88,
             size=36, bold=True, color=NAVY, align=PP_ALIGN.LEFT)
    add_text(slide, "BMAP Market Snapshot", 1.2, 2.04, 8.5, 0.34,
             size=16, color=NAVY)
    add_text(slide, d["date"], 1.2, 2.42, 8.5, 0.26, size=11, color=GRAY3)

    # 4 KPI tiles
    kpis = [
        (d["branchCount"],  "BRANCHES"),
        (d["deposits"],     "TOTAL DEPOSITS"),
        (d["avgScore"],     "AVG OPP SCORE"),
        (d["gap"],          "VS PEER AVG"),
    ]
    gap_neg = d["gapNeg"]
    for i, (val, lbl) in enumerate(kpis):
        kx = 1.2 + i*2.14
        add_rect(slide, kx, 2.82, 2.0, 0.88, GRAY1, GRAY2, Pt(0.5))
        c = JUSTIFY if (i==3 and gap_neg) else NAVY
        add_text(slide, val, kx, 2.88, 2.0, 0.48,
                 size=22, bold=True, color=c, align=PP_ALIGN.CENTER)
        add_text(slide, lbl, kx, 3.34, 2.0, 0.26,
                 size=7, bold=True, color=GRAY3, align=PP_ALIGN.CENTER)

    # 4 zone tiles
    zones = [
        (str(d["invest"]),  "INVEST",   INVEST,  INVEST_L),
        (str(d["analyze"]), "ANALYZE",  ANALYZE, ANALYZE_L),
        (str(d["defend"]),  "DEFEND",   DEFEND,  DEFEND_L),
        (str(d["justify"]), "JUSTIFY",  JUSTIFY, JUSTIFY_L),
    ]
    for i, (val, lbl, c, bg) in enumerate(zones):
        zx = 1.2 + i*2.14
        add_rect(slide, zx, 3.82, 2.0, 0.72, bg, c, Pt(0.8))
        add_text(slide, val, zx, 3.86, 2.0, 0.36, size=20, bold=True, color=c, align=PP_ALIGN.CENTER)
        add_text(slide, lbl, zx, 4.22, 2.0, 0.24, size=8,  bold=True, color=c, align=PP_ALIGN.CENTER)

    # Logo bottom-left + footer
    if logo_bytes:
        slide.shapes.add_picture(io.BytesIO(logo_bytes), Inches(0.42), Inches(5.04), Inches(1.55), Inches(0.36))
    add_text(slide,
        f"Confidential  ·  Verlocity Princeton Partners Group  ·  {datetime.now().year}",
        1.2, 5.3, 8.5, 0.2, size=7.5, color=GRAY3, align=PP_ALIGN.CENTER)


def build_network(prs, d, narr, logo_bytes):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_chrome(slide, 1, "MARKET OVERVIEW", logo_bytes)
    add_narrative(slide, narr["network"], 0.14)

    # 4 KPI tiles right
    kpis = [
        (d["deposits"],   "TOTAL DEPOSITS", GRAY1,    NAVY),
        (d["avgScore"],   "AVG OPP SCORE",  GRAY1,    NAVY),
        (d["depositYoY"], "DEPOSIT YoY",    INVEST_L if not d["gapNeg"] else JUSTIFY_L,
                                            INVEST   if not d["gapNeg"] else JUSTIFY),
        (d["gap"],        "GAP VS PEERS",   INVEST_L if not d["gapNeg"] else JUSTIFY_L,
                                            INVEST   if not d["gapNeg"] else JUSTIFY),
    ]
    for i, (val, lbl, bg, vc) in enumerate(kpis):
        kx = 6.46 + (i%2)*1.76
        ky = 0.14 + (i//2)*0.90
        add_rect(slide, kx, ky, 1.62, 0.78, bg, GRAY2, Pt(0.4))
        add_text(slide, val, kx, ky+0.06, 1.62, 0.42, size=19, bold=True, color=vc, align=PP_ALIGN.CENTER)
        add_text(slide, lbl, kx, ky+0.52, 1.62, 0.20, size=6.5, bold=True, color=GRAY3, align=PP_ALIGN.CENTER)

    # ── NATIVE VECTOR PIE CHART ──────────────────────────────────
    chart_data = ChartData()
    chart_data.categories = ["Invest", "Analyze", "Defend", "Justify"]
    chart_data.add_series("Zones", (d["invest"], d["analyze"], d["defend"], d["justify"]))

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE,
        Inches(6.0), Inches(1.8), Inches(3.8), Inches(3.1),
        chart_data
    )
    chart = chart_frame.chart

    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False

    # Color each slice
    colors = [INVEST, ANALYZE, DEFEND, JUSTIFY]
    for i, point in enumerate(chart.series[0].points):
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = colors[i]
        point.format.line.color.rgb = WHITE
        point.format.line.width = Pt(2)

    # Remove graphic frame border via XML
    from pptx.oxml.ns import qn as _qn
    from lxml import etree as _et
    sp_pr = chart_frame.element.find('.//' + _qn('c:spPr'))
    if sp_pr is None:
        sp_pr = _et.SubElement(chart_frame.element, _qn('c:spPr'))
    ln = sp_pr.find(_qn('a:ln'))
    if ln is None:
        ln = _et.SubElement(sp_pr, _qn('a:ln'))
    if ln.find(_qn('a:noFill')) is None:
        _et.SubElement(ln, _qn('a:noFill'))


def build_branches(prs, d, narr, logo_bytes):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_chrome(slide, 2, "PRIORITY MARKETS", logo_bytes)
    add_narrative(slide, narr["priority"], 0.14)

    for i, b in enumerate(d["branchList"][:5]):
        by = 0.14 + i*1.06
        zc  = ZONE_C.get(b["zone"], ANALYZE)
        zbg = ZONE_L.get(b["zone"], ANALYZE_L)

        add_rect(slide, 6.22, by, 3.6, 0.9, GRAY1, GRAY2, Pt(0.4))
        add_rect(slide, 6.22, by, 0.06, 0.9, zc)

        # Rank badge
        add_rect(slide, 6.34, by+0.26, 0.34, 0.34, zc)
        add_text(slide, str(i+1), 6.34, by+0.26, 0.34, 0.34,
                 size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        add_text(slide, b["name"], 6.76, by+0.07, 2.18, 0.26, size=10, bold=True, color=NAVY)
        add_text(slide, b["city"], 6.76, by+0.33, 2.18, 0.18, size=8, color=GRAY3)
        add_text(slide, f"{b['dep']}  ·  {b['yoy']}% YoY", 6.76, by+0.54, 2.18, 0.22, size=9, color=NAVY)

        # Zone pill
        add_rect(slide, 9.0, by+0.30, 0.72, 0.24, zbg, zc, Pt(0.5))
        add_text(slide, b["zone"], 9.0, by+0.30, 0.72, 0.24,
                 size=7, bold=True, color=zc, align=PP_ALIGN.CENTER)


def build_financial(prs, d, narr, logo_bytes):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_chrome(slide, 3, "FINANCIAL HEALTH", logo_bytes)
    add_narrative(slide, narr["financial"], 0.14)

    # Column headers
    cols = [(6.22, 1.12, "", PP_ALIGN.LEFT),
            (7.38, 1.26, "VALUE",     PP_ALIGN.CENTER),
            (8.68, 1.0,  "BENCHMARK", PP_ALIGN.CENTER)]
    for cx, cw, cl, ca in cols:
        add_rect(slide, cx, 0.14, cw, 0.30, NAVY)
        add_text(slide, cl, cx, 0.14, cw, 0.30, size=7.5, bold=True, color=WHITE, align=ca)

    for i, m in enumerate(d["metrics"]):
        my = 0.48 + i*0.64
        bg = GRAY1 if i%2==0 else WHITE
        add_rect(slide, 6.22, my, 3.56, 0.58, bg, GRAY2, Pt(0.3))
        add_text(slide, m["label"], 6.30, my+0.13, 1.0, 0.28, size=9.5, color=NAVY)
        add_text(slide, m["value"], 7.38, my+0.10, 1.26, 0.32, size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        add_text(slide, m["bench"], 8.68, my+0.13, 1.0,  0.28, size=9, italic=True, color=GRAY3, align=PP_ALIGN.CENTER)

        sc = TEAL if m["ok"] else AMBER
        add_rect(slide, 9.74, my+0.13, 0.30, 0.30, sc)
        add_text(slide, "✓" if m["ok"] else "!", 9.74, my+0.13, 0.30, 0.30,
                 size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    if d.get("competitor"):
        add_rect(slide, 6.22, 4.88, 3.56, 0.34, JUSTIFY_L, JUSTIFY, Pt(0.5))
        add_text(slide,
            f"⚠  Key Competitor  ·  {d['competitor']['branches']} branch overlap"
            f"  ·  Peer avg YoY {d['competitor']['yoy']}%",
            6.32, 4.90, 3.36, 0.28, size=8.5, bold=True, color=JUSTIFY)


def build_gap(prs, d, narr):
    """Dark slide — no chrome, uses navy background"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Dark background
    bg = slide.background; fill = bg.fill; fill.solid(); fill.fore_color.rgb = NAVY

    # Left teal stripe only
    add_rect(slide, 0, 0, 0.12, 5.625, TEAL)

    # Giant gap number
    add_text(slide, d["gap"], 0.28, 0.16, 5.2, 1.86,
             size=96, bold=True, color=TEAL, align=PP_ALIGN.LEFT)
    add_text(slide, "GAP VS PEER AVERAGE", 0.28, 2.08, 5.2, 0.34,
             size=13, bold=True, color=WHITE)
    add_text(slide, d["gapSubtitle"], 0.28, 2.50, 5.2, 0.26,
             size=9.5, italic=True, color=rgb("4A6A8A"))

    # 3 stat tiles
    tile_c = rgb("F87171") if d["gapNeg"] else TEAL
    tiles = [
        (f"{d['bankYoY']}%", "THIS BANK YoY", tile_c),
        (f"{d['peerYoY']}%", "PEER AVG",      GRAY3),
        (d["gap"],           "GAP",            AMBER),
    ]
    for i, (val, lbl, c) in enumerate(tiles):
        tx = 0.28 + i*1.78
        add_rect(slide, tx, 2.96, 1.62, 1.12, rgb("162436"), rgb("1E3A5F"), Pt(0.5))
        add_rect(slide, tx, 2.96, 1.62, 0.06, c)
        add_text(slide, val, tx, 3.06, 1.62, 0.58, size=20, bold=True, color=c, align=PP_ALIGN.CENTER)
        add_text(slide, lbl, tx, 3.70, 1.62, 0.26, size=7,  bold=True, color=rgb("3A5A7A"), align=PP_ALIGN.CENTER)

    # ── NATIVE VECTOR BAR CHART ──────────────────────────────────
    chart_data = ChartData()
    chart_data.categories = ["This Bank", "Peer Avg"]
    chart_data.add_series("Deposit YoY %", (float(d["bankYoY"]), float(d["peerYoY"])))

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(5.4), Inches(0.15), Inches(4.4), Inches(5.0),
        chart_data
    )
    chart = chart_frame.chart

    chart.has_legend = False
    chart.has_title  = False

    # Bar colors
    bar_colors = [rgb("A32D2D") if d["gapNeg"] else TEAL, GRAY3]
    for i, point in enumerate(chart.series[0].points):
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = bar_colors[i]
        point.format.line.fill.background()

    # Style axes
    va = chart.value_axis
    va.tick_labels.font.color.rgb = GRAY3
    va.tick_labels.font.size = Pt(11)

    ca = chart.category_axis
    ca.tick_labels.font.color.rgb = WHITE
    ca.tick_labels.font.size = Pt(13)
    ca.tick_labels.font.bold = True

    # Footer
    add_text(slide,
        f"Verlocity Princeton Partners Group   ·   BMAP Intelligence   ·   {d['bankName']}",
        0.28, 5.30, 9.5, 0.22, size=7.5, color=rgb("2A4060"))
    add_text(slide, "5", 9.50, 5.30, 0.38, 0.22, size=9, color=rgb("2A4060"), align=PP_ALIGN.RIGHT)


def build_next_steps(prs, d, narr, logo_bytes):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_chrome(slide, 5, "STRATEGIC PRIORITIES", logo_bytes)
    add_narrative(slide, narr["nextsteps"], 0.14)

    ac_colors = [TEAL, ANALYZE, AMBER, NAVY]
    for i, action in enumerate(d["actions"]):
        ay = 0.14 + i*1.30
        ac = ac_colors[i]

        add_rect(slide, 6.22, ay, 3.6, 1.14, GRAY1, GRAY2, Pt(0.4))
        add_rect(slide, 6.22, ay, 0.06, 1.14, ac)
        add_rect(slide, 6.34, ay+0.36, 0.34, 0.34, ac)
        add_text(slide, str(i+1).zfill(2), 6.34, ay+0.36, 0.34, 0.34,
                 size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(slide, action["title"], 6.76, ay+0.10, 2.98, 0.28,
                 size=10.5, bold=True, color=NAVY)
        add_text(slide, action["body"],  6.76, ay+0.42, 2.98, 0.62,
                 size=8.5, color=GRAY3)


# ═══════════════════════════════════════════════════════════════
# MAIN BUILD
# ═══════════════════════════════════════════════════════════════

def build_deck(data, logo_bytes):
    rows = data["rows"]; fin = data["fin"]; tgt = data["tgt"]
    br   = data["br"];   bankName = data["bankName"]

    def sf(v, default=0):
        """Safe float — handles None, empty string, missing."""
        try:
            return float(v) if v is not None else default
        except (TypeError, ValueError):
            return default

    tot      = sum(sf(r.get("latest_dep")) for r in rows)
    avg      = lambda v: sum(sf(r.get(v)) for r in rows) / max(len(rows),1)
    invest   = sum(1 for r in rows if r.get("opportunity_zone")=="Invest")
    analyze  = sum(1 for r in rows if r.get("opportunity_zone")=="Analyze")
    defend   = sum(1 for r in rows if r.get("opportunity_zone")=="Defend")
    justify  = sum(1 for r in rows if r.get("opportunity_zone")=="Justify")
    bankYoY  = avg("yoy_deposits")*100
    compYoY  = avg("avg_comp_yoy")*100
    gap      = bankYoY - compYoY
    avgScore = avg("opportunity_score")

    top_br   = sorted(br, key=lambda b: sf(b.get("opportunity_score")), reverse=True)
    just_top = sorted([b for b in br if b.get("opportunity_zone")=="Justify"],
                      key=lambda b: sf(b.get("latest_dep")), reverse=True)
    tier1    = [b for b in br if (b.get("priority_tier") or "").startswith("1")][:2]

    narr = get_narratives(data)

    D = {
        "bankName":    bankName,
        "date":        datetime.now().strftime("%B %Y"),
        "branchCount": str(len(rows)),
        "deposits":    f"${tot/1e9:.1f}B",
        "avgScore":    f"{avgScore:.0f}",
        "depositYoY":  f"{bankYoY:+.1f}%",
        "gap":         f"{gap:+.1f}pp",
        "bankYoY":     f"{bankYoY:.1f}",
        "peerYoY":     f"{compYoY:.1f}",
        "gapNeg":      gap < 0,
        "gapSubtitle": f"Deposit growth vs. peer average — {fin.get('period','Q4 2025')}",
        "invest":  invest,  "analyze": analyze,
        "defend":  defend,  "justify": justify,
        "branchList": [
            {
                "name":  b["namebr"].split("--")[-1].strip()[:16],
                "city":  f"{b['citybr']}, {b['stalpbr']}",
                "score": f"{sf(b.get('opportunity_score')):.0f}",
                "dep":   f"${sf(b.get('latest_dep'))/1e6:.0f}M",
                "yoy":   f"{sf(b.get('yoy_deposits'))*100:+.1f}",
                "zone":  b.get("opportunity_zone",""),
            }
            for b in top_br[:8]
        ],
        "metrics": [
            {"label":"ROA",           "value":f"{fin.get('roa','—')}%",              "bench":">1.0%",    "ok": sf(fin.get("roa"))>=1},
            {"label":"NIM",           "value":f"{fin.get('nim','—')}%",              "bench":"2.5–3.5%", "ok": 2.5<=sf(fin.get("nim"))<=4},
            {"label":"Efficiency",    "value":f"{fin.get('efficiency_ratio','—')}%", "bench":"<60%",     "ok": 0<sf(fin.get("efficiency_ratio"))<60},
            {"label":"Net Income YoY","value":f"{sf(fin.get('net_income_yoy_pct')):+.1f}%",              "bench":">0%",      "ok": sf(fin.get("net_income_yoy_pct"))>0},
            {"label":"Deposit YoY",   "value":f"{bankYoY:+.1f}%",                   "bench":">2%",      "ok": bankYoY>=2},
            {"label":"Cost of Funds", "value":f"{fin.get('cost_of_funds_pct','—')}%","bench":"<2%",      "ok": 0<sf(fin.get("cost_of_funds_pct"))<2},
            {"label":"Tier 1 Capital","value":f"{fin.get('tier1_capital_pct','—')}%","bench":">8%",      "ok": sf(fin.get("tier1_capital_pct"))>=8},
        ],
        "competitor": {
            "branches": tgt["branches_in_radius"],
            "yoy":      f"{sf(tgt.get('avg_yoy_pct')):.1f}"
        } if tgt else None,
        "actions": [
            {
                "title": f"Activate: {' + '.join(' '.join(b['namebr'].split('--')[-1].strip().split()[:3]) for b in tier1)}" if tier1 else "Activate Top Invest Branches",
                "body":  "\n".join(f"{b['namebr'].split('--')[-1].strip()} — Score {sf(b.get('opportunity_score')):.0f} | ${sf(b.get('latest_dep'))/1e6:.0f}M" for b in tier1) or f"{invest} Invest zone branches ready for campaign launch.",
            },
            {
                "title": "Launch Targeted Audience Campaigns",
                "body":  f"{invest} Invest zone branches. Rate-sensitive depositors + digital big-bank leavers. Deploy via AudienceFinder.",
            },
            {
                "title": f"Justify Zone — {justify} Branches Under Review",
                "body":  "\n".join(f"{b['namebr'].split('--')[-1].strip()}: ${sf(b.get('latest_dep'))/1e6:.0f}M — assess ROI" for b in just_top[:2]) or f"{justify} branches need investment audit.",
            },
            {
                "title": f"Protect Against {tgt['target_institution']}" if tgt else "Protect Market Position",
                "body":  f"{tgt['branches_in_radius']} shared geographies. Competitor at {sf(tgt.get('avg_yoy_pct')):.1f}% YoY. Deploy defensive rate messaging." if tgt else "Identify top competitors and monitor rate activity.",
            },
        ],
    }

    print(f"  Building slides...")
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H
    # Ensure blank layout exists
    while len(prs.slide_layouts) < 7:
        prs.slide_layouts.add_slide()

    build_cover(prs, D, logo_bytes)
    build_network(prs, D, narr, logo_bytes)
    build_branches(prs, D, narr, logo_bytes)
    build_financial(prs, D, narr, logo_bytes)
    build_gap(prs, D, narr)
    build_next_steps(prs, D, narr, logo_bytes)

    return prs


def save_deck(prs, bank_name, out_dir=OUT_DIR):
    safe = "".join(c if c.isalnum() or c in " _-" else "_" for c in bank_name).strip()
    date = datetime.now().strftime("%Y%m%d")
    fname = out_dir / f"BMAP_Snapshot_{safe}_{date}.pptx"
    prs.save(str(fname))
    return fname


# ═══════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════

def run_single(ik, name_hint=None):
    print(f"\n{'='*55}")
    print(f"  BMAP Snapshot — {name_hint or ik}")
    print(f"{'='*55}")
    data = fetch_bank_data(ik)
    if name_hint:
        data["bankName"] = name_hint
    logo = fetch_logo()
    prs  = build_deck(data, logo)
    path = save_deck(prs, data["bankName"])
    print(f"\n  ✓  Saved: {path}\n")
    return path


def run_batch(csv_path):
    banks = []
    with open(csv_path, newline="") as f:
        for row in csv.reader(f):
            if not row or row[0].startswith("#") or row[0].lower()=="inst_key":
                continue
            ik   = row[0].strip()
            name = row[1].strip() if len(row)>1 else ""
            banks.append((ik, name))

    print(f"\n  Batch: {len(banks)} banks from {csv_path}")
    logo = fetch_logo()
    out_dir = OUT_DIR / f"BMAP_Batch_{datetime.now().strftime('%Y%m%d_%H%M')}"
    out_dir.mkdir(parents=True, exist_ok=True)

    results = []
    for i, (ik, name) in enumerate(banks):
        print(f"\n[{i+1}/{len(banks)}] {name or ik}")
        try:
            data = fetch_bank_data(ik)
            if name: data["bankName"] = name
            prs  = build_deck(data, logo)
            path = save_deck(prs, data["bankName"], out_dir)
            print(f"  ✓  {path.name}")
            results.append({"bank": data["bankName"], "status": "ok", "file": str(path)})
        except Exception as e:
            print(f"  ✗  FAILED: {e}")
            results.append({"bank": name or ik, "status": "error", "error": str(e)})
        if i < len(banks)-1:
            time.sleep(0.5)

    ok  = sum(1 for r in results if r["status"]=="ok")
    err = sum(1 for r in results if r["status"]=="error")
    print(f"\n{'='*55}")
    print(f"  Batch complete — {ok} decks saved, {err} failed")
    print(f"  Output folder: {out_dir}")
    print(f"{'='*55}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="BMAP Snapshot Deck Builder — generates branded PPTX from Supabase data"
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--inst_key", help="Single bank inst_key (e.g. bank_123)")
    group.add_argument("--csv",      help="CSV file with inst_key[,name] per line")
    parser.add_argument("--name",    help="Override bank name (single mode only)")
    parser.add_argument("--out",     help="Output directory (default: current folder)", default=".")
    args = parser.parse_args()

    OUT_DIR = Path(args.out)
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    if not ANTH_KEY:
        print("\n  ⚠  Set ANTHROPIC_API_KEY env var for AI narratives.")
        print("     export ANTHROPIC_API_KEY=sk-ant-...")
        print("     Continuing with empty narratives...\n")

    if args.inst_key:
        run_single(args.inst_key, args.name)
    else:
        run_batch(args.csv)
