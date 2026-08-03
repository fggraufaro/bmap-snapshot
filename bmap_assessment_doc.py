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

# ── Brand colors ────────────────────────────────────────────────
NAVY  = RGBColor(0x1A, 0x23, 0x32)
TEAL  = RGBColor(0x1D, 0x9E, 0x75)
AMBER = RGBColor(0xF5, 0xA6, 0x23)
GRAY3 = RGBColor(0x8A, 0x8A, 0x80)
GRAY_FILL = "F7F7F5"
ZONE_COLOR = {
    "Invest":  "1A2332",
    "Analyze": "1D9E75",
    "Defend":  "8A8A80",
    "Justify": "C0392B",
}

# ═══════════════════════════════════════════════════════════════
# DATA FETCH — full network, no truncation
# ═══════════════════════════════════════════════════════════════

# SCHEMA_MAP mirrors bmap_snapshot.py exactly — table→schema routing,
# since PostgREST needs Accept-Profile for anything outside 'public'.
SCHEMA_MAP = {
    "branch_opportunity_base":        "analytics",
    "bank_financial_snapshot_latest": "analytics",
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


def fetch_full_network_data(ik):
    """Pull FULL branch network — no limit=N slice. This is the structural
    difference vs. the free Snapshot."""
    print(f"  Fetching full branch network for {ik}...")
    branches = supabase(
        "branch_opportunity_base",
        f"inst_key=eq.{ik}&select=uninumbr,namebr,citybr,stalpbr,latest_dep,"
        "yoy_deposits,opportunity_score,opportunity_zone,matrix_quadrant,"
        "priority_tier,market_growth_score,rel_growth_norm,namefull&order=opportunity_score.desc",
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

    return {
        "inst_key": ik,
        "branches": branches,
        "fin": fin,
        "targets": tgt_arr,
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
    try:
        return float(v) if v is not None else default
    except (TypeError, ValueError):
        return default


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

def get_narratives(bank_name, summary, fin, targets):
    if not ANTH_KEY or not anthropic:
        print("  ⚠ No ANTHROPIC_API_KEY — using placeholder narratives")
        return _placeholder_narratives()

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
"""

    system = """You are writing the $10K Verlocity BMAP Assessment for a community bank CFO/CEO audience.
Tone: precise, CFO-appropriate. No superlatives, no urgency language, no self-referential commentary.
State facts, name specific branches/competitors, quantify everything possible.
Return ONLY valid JSON, no markdown fences:
{
  "exec_summary": "3-4 sentences. The single most important finding, stated plainly, with a number.",
  "network_narrative": "2-3 sentences on what the zone distribution reveals about the network's overall position.",
  "competitive_narrative": "2-3 sentences naming the specific network-level target and why it is vulnerable.",
  "financial_narrative": "2-3 sentences on what the financial metrics mean together — not a list restated as prose.",
  "next_step": "2-3 sentences. A specific, named recommendation tied to the top opportunity branches."
}"""

    print("  Generating AI narratives (full-network context)...")
    client = anthropic.Anthropic(api_key=ANTH_KEY)
    try:
        msg = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1200,
            system=system,
            messages=[{"role": "user", "content": ctx}],
        )
        raw = msg.content[0].text.strip().replace("```json", "").replace("```", "").strip()
        narr = json.loads(raw)
        print("  ✓ Narratives generated")
        return narr
    except Exception as e:
        print(f"  ⚠ Narrative generation failed ({e}) — using placeholders")
        return _placeholder_narratives()


def _placeholder_narratives():
    return {k: "" for k in ["exec_summary", "network_narrative", "competitive_narrative",
                             "financial_narrative", "next_step"]}


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
    return p


def _body(doc, text, size=10.5, color=RGBColor(0x33, 0x33, 0x33)):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.color.rgb = color
    return p


def build_assessment_doc(bank_name, summary, fin, targets, narr, branches):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.59)   # US Letter
    section.page_height = Cm(27.94)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)

    # ── Cover ──
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(60)
    run = p.add_run("VERLOCITY")
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = TEAL

    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(6)
    run2 = p2.add_run(bank_name)
    run2.bold = True
    run2.font.size = Pt(28)
    run2.font.color.rgb = NAVY

    p3 = doc.add_paragraph()
    run3 = p3.add_run("BMAP Market Assessment")
    run3.font.size = Pt(14)
    run3.font.color.rgb = GRAY3

    p4 = doc.add_paragraph()
    p4.paragraph_format.space_after = Pt(40)
    run4 = p4.add_run(datetime.now().strftime("%B %Y"))
    run4.font.size = Pt(10)
    run4.font.color.rgb = GRAY3

    doc.add_page_break()

    # ── Executive Summary ──
    _heading(doc, "Executive Summary", space_before=0)
    _body(doc, narr.get("exec_summary") or
          f"{bank_name} operates {summary['branch_count']} branches with ${summary['total_deposits_B']:.1f}B "
          f"in total deposits. Network average opportunity score: {summary['avg_score']:.0f}/100.")

    # ── Network Opportunity Overview ──
    _heading(doc, "Network Opportunity Overview")
    _body(doc, narr.get("network_narrative") or "")

    zt = doc.add_table(rows=2, cols=4)
    zt.alignment = WD_TABLE_ALIGNMENT.LEFT
    zones = summary["zones"]
    for i, z in enumerate(["Invest", "Analyze", "Defend", "Justify"]):
        c0 = zt.cell(0, i)
        c0.text = z
        _set_cell_shading(c0, ZONE_COLOR[z])
        for run in c0.paragraphs[0].runs:
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.bold = True
        c1 = zt.cell(1, i)
        c1.text = str(zones[z])
        c1.paragraphs[0].runs[0].font.size = Pt(16)
        c1.paragraphs[0].runs[0].bold = True
    doc.add_paragraph().paragraph_format.space_after = Pt(10)

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

    # ── Financial Health Benchmarking ──
    _heading(doc, "Financial Health Benchmarking")
    _body(doc, narr.get("financial_narrative") or "")
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
    narr = get_narratives(bank_name, summary, d["fin"], d["targets"])
    doc = build_assessment_doc(bank_name, summary, d["fin"], d["targets"], narr, d["branches"])
    path = save_doc(doc, bank_name)
    print(f"\n  ✓ Saved: {path}\n")
    return path


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--inst_key", required=True)
    ap.add_argument("--name", default=None)
    args = ap.parse_args()
    run(args.inst_key, args.name)
