"""
bmap_board_brief.py — Verlocity Board Brief PDF Generator
==========================================================
Generates a 5-page executive board brief for a target bank.
Tone: CFO-written, not agency-written. Numbers first. No jargon.
Uses the same Supabase data as bmap_snapshot.py.

Usage (CLI):
    python bmap_board_brief.py --inst_key bank_463735
    python bmap_board_brief.py --inst_key bank_463735 --name "Hancock Whitney Bank"

Railway API: called from main.py via generate_board_brief()
"""

import io
import os
import sys
import csv
import requests
from datetime import datetime
from pathlib import Path

# ── ReportLab imports ─────────────────────────────────────────
from reportlab.lib.pagesizes   import landscape, letter
from reportlab.lib.units       import inch
from reportlab.lib             import colors
from reportlab.lib.styles      import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums       import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus        import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether, PageBreak
)
from reportlab.graphics.shapes import Drawing, Rect, String, Line
from reportlab.graphics        import renderPDF

# ── Anthropic ─────────────────────────────────────────────────
try:
    import anthropic as _anthropic
except ImportError:
    _anthropic = None

import json

ANTH_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

# ── Supabase config (same as bmap_snapshot.py) ────────────────
SUPA_URL = "https://tuiiywphoynbmkxpoyps.supabase.co"
SUPA_KEY = os.environ.get(
    "SUPABASE_KEY",
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6"
    "InR1aWl5d3Bob3luYm1reHBveXBzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NTc0MDg0NT"
    "MsImV4cCI6MjA3Mjk4NDQ1M30.8-JAz4WQRE3Fi6uH7xiYNTns92g-nV1A9pbUvSK549M"
)
LOGO_URL = "https://fggraufaro.github.io/bmap-tools/Verlocity-Logo.png"
OUT_DIR  = Path(".")

# ── Brand colors ──────────────────────────────────────────────
NAVY     = colors.HexColor("#1A2332")
TEAL     = colors.HexColor("#1D9E75")
AMBER    = colors.HexColor("#F5A623")
GRAY1    = colors.HexColor("#F7F7F5")
GRAY2    = colors.HexColor("#E8E8E4")
GRAY3    = colors.HexColor("#8A8A80")
WHITE    = colors.white
RED_SOFT = colors.HexColor("#C0392B")

# ── Page setup (landscape letter) ─────────────────────────────
PAGE_W, PAGE_H = landscape(letter)   # 11 × 8.5 inches
MARGIN = 0.65 * inch

# ══════════════════════════════════════════════════════════════
# DATA FETCH
# ══════════════════════════════════════════════════════════════

def _supa(table, params):
    url = f"{SUPA_URL}/rest/v1/{table}?{params}"
    r = requests.get(url,
        headers={"apikey": SUPA_KEY, "Authorization": f"Bearer {SUPA_KEY}"},
        timeout=25)
    r.raise_for_status()
    return r.json()


def _sf(v, default=0):
    try:
        return float(v) if v is not None else default
    except (TypeError, ValueError):
        return default


def fetch_board_data(ik):
    """Fetch all data needed for the board brief."""
    print(f"  [board] Fetching branch data...")
    rows = _supa("branch_opportunity_base",
        f"inst_key=eq.{ik}&select=namefull,latest_dep,yoy_deposits,"
        "avg_comp_yoy,opportunity_zone,opportunity_score")

    print(f"  [board] Fetching financials...")
    fin_arr = _supa("bank_financial_snapshot_latest",
        f"inst_key=eq.{ik}&select=*&limit=1")
    fin = fin_arr[0] if fin_arr else {}

    print(f"  [board] Fetching competitor data...")
    tgt_arr = _supa("vw_network_top_targets",
        f"my_inst_key=eq.{ik}&select=target_institution,branches_in_radius,"
        "avg_vuln_score,avg_yoy_pct&order=network_rank.asc&limit=1")
    tgt = tgt_arr[0] if tgt_arr else None

    print(f"  [board] Fetching brokered deposits...")
    rssdid = ik.replace("bank_", "").replace("cu_", "")
    brok = None
    try:
        rce = requests.get(
            f"{SUPA_URL}/rest/v1/raw_schedule_RCE"
            f"?IDRSSD=eq.{rssdid}&period=eq.2025-12-31"
            f"&select=RCON2365,RCON2385&limit=1",
            headers={"apikey": SUPA_KEY, "Authorization": f"Bearer {SUPA_KEY}"},
            timeout=15
        ).json()
        if rce:
            brokered  = _sf(rce[0].get("RCON2365"))
            total_dep = _sf(rce[0].get("RCON2385"))
            if total_dep > 0:
                pct = brokered / total_dep
                if pct >= 0.15:
                    brok = {
                        "pct":        round(pct * 100, 1),
                        "brokered_M": round(brokered / 1000, 1),
                        "total_M":    round(total_dep / 1000, 1),
                    }
    except Exception as e:
        print(f"  [board] brokered fetch error: {e}")

    bank_name = rows[0].get("namefull", ik) if rows else ik
    tot       = sum(_sf(r.get("latest_dep")) for r in rows)
    avg_yoy   = lambda v: sum(_sf(r.get(v)) for r in rows) / max(len(rows), 1)
    bank_yoy  = avg_yoy("yoy_deposits") * 100
    comp_yoy  = avg_yoy("avg_comp_yoy") * 100
    gap       = bank_yoy - comp_yoy
    invest    = sum(1 for r in rows if r.get("opportunity_zone") == "Invest")
    analyze   = sum(1 for r in rows if r.get("opportunity_zone") == "Analyze")
    defend    = sum(1 for r in rows if r.get("opportunity_zone") == "Defend")
    justify   = sum(1 for r in rows if r.get("opportunity_zone") == "Justify")
    at_risk   = defend + justify

    # Estimate deposit volume at risk: avg deposit per branch × at_risk branches
    avg_dep_per_branch = (tot / max(len(rows), 1)) if rows else 0
    vol_at_risk = avg_dep_per_branch * at_risk

    # 12-month trajectory: if gap continues, what's lost?
    projected_loss = tot * abs(gap / 100) if gap < 0 else 0

    return {
        "ik":             ik,
        "bank_name":      bank_name,
        "date":           datetime.now().strftime("%B %Y"),
        "branches":       len(rows),
        "deposits_B":     round(tot / 1e9, 1),
        "deposits_str":   f"${tot/1e9:.1f}B",
        "bank_yoy":       round(bank_yoy, 1),
        "comp_yoy":       round(comp_yoy, 1),
        "gap":            round(gap, 1),
        "gap_str":        f"{gap:+.1f}pp",
        "invest":         invest,
        "analyze":        analyze,
        "defend":         defend,
        "justify":        justify,
        "at_risk":        at_risk,
        "vol_at_risk_M":  round(vol_at_risk / 1e6),
        "proj_loss_M":    round(projected_loss / 1e6),
        "fin":            fin,
        "tgt":            tgt,
        "brok":           brok,
    }


# ══════════════════════════════════════════════════════════════
# STYLES
# ══════════════════════════════════════════════════════════════

def make_styles():
    base = getSampleStyleSheet()

    def s(name, **kw):
        return ParagraphStyle(name, **kw)

    return {
        "cover_bank": s("cover_bank",
            fontName="Helvetica-Bold", fontSize=28,
            textColor=NAVY, leading=34, spaceAfter=6),

        "cover_sub": s("cover_sub",
            fontName="Helvetica", fontSize=13,
            textColor=GRAY3, leading=18, spaceAfter=4),

        "cover_date": s("cover_date",
            fontName="Helvetica", fontSize=11,
            textColor=GRAY3, leading=16),

        "section_label": s("section_label",
            fontName="Helvetica-Bold", fontSize=7.5,
            textColor=TEAL, leading=10, spaceBefore=18, spaceAfter=4,
            wordWrap='LTR'),

        "page_headline": s("page_headline",
            fontName="Helvetica-Bold", fontSize=18,
            textColor=NAVY, leading=22, spaceAfter=8),

        "body": s("body",
            fontName="Helvetica", fontSize=10.5,
            textColor=NAVY, leading=16, spaceAfter=6, alignment=TA_JUSTIFY),

        "body_italic": s("body_italic",
            fontName="Helvetica-Oblique", fontSize=10,
            textColor=GRAY3, leading=15, spaceAfter=6, alignment=TA_JUSTIFY),

        "callout": s("callout",
            fontName="Helvetica-Bold", fontSize=11,
            textColor=NAVY, leading=16, spaceAfter=4, alignment=TA_JUSTIFY),

        "number_big": s("number_big",
            fontName="Helvetica-Bold", fontSize=38,
            textColor=NAVY, leading=42, spaceAfter=2, alignment=TA_CENTER),

        "number_label": s("number_label",
            fontName="Helvetica-Bold", fontSize=7,
            textColor=GRAY3, leading=10, spaceAfter=0, alignment=TA_CENTER,
            wordWrap='LTR'),

        "question": s("question",
            fontName="Helvetica-Bold", fontSize=11,
            textColor=NAVY, leading=16, spaceAfter=2),

        "question_body": s("question_body",
            fontName="Helvetica-Oblique", fontSize=10,
            textColor=GRAY3, leading=14, spaceAfter=10, alignment=TA_JUSTIFY),

        "footer": s("footer",
            fontName="Helvetica", fontSize=7.5,
            textColor=GRAY3, leading=10, alignment=TA_CENTER),

        "ask_label": s("ask_label",
            fontName="Helvetica-Bold", fontSize=9,
            textColor=GRAY3, leading=12, spaceAfter=2),

        "ask_value": s("ask_value",
            fontName="Helvetica-Bold", fontSize=12,
            textColor=NAVY, leading=16, spaceAfter=8),

        "brokered_label": s("brokered_label",
            fontName="Helvetica-Bold", fontSize=8,
            textColor=AMBER, leading=10, spaceAfter=2),

        "brokered_body": s("brokered_body",
            fontName="Helvetica", fontSize=9.5,
            textColor=NAVY, leading=14, spaceAfter=4, alignment=TA_JUSTIFY),
    }


# ══════════════════════════════════════════════════════════════
# PAGE TEMPLATE
# ══════════════════════════════════════════════════════════════

def make_page_template(bank_name, page_label):
    """Returns onFirstPage / onLaterPages callback for this page."""

    def draw_chrome(canvas, doc):
        canvas.saveState()

        # Navy left stripe
        canvas.setFillColor(NAVY)
        canvas.rect(0, 0, 0.18 * inch, PAGE_H, fill=1, stroke=0)

        # Teal top rule
        canvas.setFillColor(TEAL)
        canvas.rect(0.18 * inch, PAGE_H - 0.06 * inch,
                    PAGE_W - 0.18 * inch, 0.06 * inch, fill=1, stroke=0)

        # Verlocity wordmark bottom left
        canvas.setFillColor(NAVY)
        canvas.setFont("Helvetica-Bold", 9)
        canvas.drawString(0.32 * inch, 0.28 * inch, "Verlocity")
        canvas.setFillColor(TEAL)
        canvas.setFont("Helvetica-Bold", 9)
        canvas.drawString(0.32 * inch + canvas.stringWidth("Verlocity", "Helvetica-Bold", 9) + 2,
                          0.28 * inch, "▲")

        # Footer center
        canvas.setFillColor(GRAY3)
        canvas.setFont("Helvetica", 7.5)
        footer = f"Confidential  ·  Verlocity Princeton Partners Group  ·  {bank_name}  ·  {page_label}"
        canvas.drawCentredString(PAGE_W / 2, 0.28 * inch, footer)

        # Page number bottom right
        canvas.setFont("Helvetica", 7.5)
        canvas.drawRightString(PAGE_W - 0.4 * inch, 0.28 * inch,
                               str(doc.page))

        canvas.restoreState()

    return draw_chrome


# ══════════════════════════════════════════════════════════════
# NUMBER TILE HELPER
# ══════════════════════════════════════════════════════════════

def num_tile(value_str, label, color=None, bg=GRAY1):
    """Returns a single-cell table acting as a KPI tile."""
    c = color or NAVY
    tbl = Table(
        [[Paragraph(f'<font color="#{c.hexval()[2:]}" size="26"><b>{value_str}</b></font>', ParagraphStyle("v", alignment=TA_CENTER, leading=30))],
         [Paragraph(f'<font color="#{GRAY3.hexval()[2:]}" size="7"><b>{label}</b></font>', ParagraphStyle("l", alignment=TA_CENTER, leading=10))]],
        colWidths=[1.8 * inch],
        rowHeights=[0.48 * inch, 0.22 * inch],
    )
    tbl.setStyle(TableStyle([
        ("BACKGROUND",  (0, 0), (-1, -1), bg),
        ("TOPPADDING",  (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING",  (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("ROUNDEDCORNERS", (0, 0), (-1, -1), [4, 4, 4, 4]),
    ]))
    return tbl


# ══════════════════════════════════════════════════════════════
# PAGE BUILDERS
# ══════════════════════════════════════════════════════════════

def page_cover(d, ST):
    """Page 1 — Cover."""
    story = []
    story.append(Spacer(1, 1.2 * inch))
    story.append(Paragraph("BOARD STRATEGIC BRIEF", ST["section_label"]))
    story.append(Paragraph(d["bank_name"], ST["cover_bank"]))
    story.append(Paragraph("Deposit Market Intelligence", ST["cover_sub"]))
    story.append(Paragraph(d["date"], ST["cover_date"]))
    story.append(Spacer(1, 0.5 * inch))

    # KPI row
    gap_color = RED_SOFT if d["gap"] < 0 else TEAL
    tiles = [
        num_tile(d["deposits_str"],         "TOTAL DEPOSITS"),
        num_tile(str(d["branches"]),         "BRANCHES ANALYZED"),
        num_tile(f"{d['bank_yoy']:+.1f}%",  "DEPOSIT YoY",
                 color=RED_SOFT if d["bank_yoy"] < 2 else TEAL),
        num_tile(d["gap_str"],               "GAP VS MKT PEERS",
                 color=gap_color),
    ]
    row_tbl = Table([tiles],
        colWidths=[2.0 * inch] * 4,
        spaceBefore=0)
    row_tbl.setStyle(TableStyle([
        ("LEFTPADDING",  (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(row_tbl)
    story.append(Spacer(1, 0.4 * inch))

    story.append(Paragraph(
        "This brief is prepared exclusively for board review. All data sourced from "
        "FDIC call reports and Verlocity's proprietary market intelligence platform.",
        ST["body_italic"]))
    story.append(PageBreak())
    return story


def page_situation(d, ST, narr={}):
    """Page 2 — The situation."""
    story = []
    story.append(Paragraph("THE SITUATION", ST["section_label"]))
    story.append(Paragraph(
        f"Deposit competition in {d['bank_name']}'s markets has intensified.",
        ST["page_headline"]))

    story.append(HRFlowable(width="100%", thickness=0.5,
                             color=TEAL, spaceAfter=12))

    # Two-column: gap chart left, narrative right
    gap_color_hex = "C0392B" if d["gap"] < 0 else "1D9E75"
    peer_num = Paragraph(
        f'<font size="11" color="#8A8A80">Peer avg</font><br/>'
        f'<font size="30" color="#1A2332"><b>{d["comp_yoy"]:+.1f}%</b></font>',
        ParagraphStyle("p", alignment=TA_CENTER, leading=36))
    bank_num = Paragraph(
        f'<font size="11" color="#8A8A80">{d["bank_name"].split()[0]} YoY</font><br/>'
        f'<font size="30" color="#{gap_color_hex}"><b>{d["bank_yoy"]:+.1f}%</b></font>',
        ParagraphStyle("p", alignment=TA_CENTER, leading=36))
    gap_num  = Paragraph(
        f'<font size="11" color="#8A8A80">Gap</font><br/>'
        f'<font size="30" color="#{gap_color_hex}"><b>{d["gap_str"]}</b></font>',
        ParagraphStyle("p", alignment=TA_CENTER, leading=36))

    chart_tbl = Table(
        [[bank_num, gap_num, peer_num]],
        colWidths=[1.7 * inch, 1.7 * inch, 1.7 * inch],
    )
    chart_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (0, 0), colors.HexColor("#FEF2F2")),
        ("BACKGROUND",    (1, 0), (1, 0), colors.HexColor("#FFF8EC")),
        ("BACKGROUND",    (2, 0), (2, 0), GRAY1),
        ("TOPPADDING",    (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 8),
        ("LINEAFTER",     (0, 0), (1, 0), 0.5, GRAY2),
    ]))

    _def_open = (
        f"{d['bank_name']} deposit growth of {d['bank_yoy']:+.1f}% "
        f"trails the competitor average in your own markets by "
        f"{abs(d['gap']):.1f} percentage points. "
        f"Peer banks operating in the same geographies grew at "
        f"{d['comp_yoy']:+.1f}% over the same period."
    )
    _def_insight = (
        "This gap is not a market condition — it is a strategy and execution "
        "gap. The same markets produced different outcomes for different "
        "institutions. Understanding why requires branch-level intelligence "
        "that aggregate reporting cannot provide."
    )
    opening   = narr.get("opening",   _def_open)
    insight   = narr.get("insight",   _def_insight)
    impl      = narr.get("implication", "")
    close_txt = narr.get("close", "")

    narrative = [
        Paragraph(opening,  ST["body"]),
        Spacer(1, 6),
        Paragraph(insight,  ST["body"]),
    ]
    if impl:
        narrative.append(Spacer(1, 4))
        narrative.append(Paragraph(impl, ST["body_italic"]))
    if close_txt:
        narrative.append(Spacer(1, 6))
        narrative.append(Paragraph(f"<b>{close_txt}</b>", ST["callout"]))

    if d["brok"]:
        narrative.append(Spacer(1, 6))
        narrative.append(Paragraph("BROKERED DEPOSIT EXPOSURE", ST["brokered_label"]))
        narrative.append(Paragraph(
            f"{d['brok']['pct']}% of total deposits (${d['brok']['brokered_M']:.0f}M) "
            f"are currently sourced through brokers — expensive, rate-sensitive capital "
            f"that leaves when a better rate appears. Converting a portion of this to "
            f"direct customer deposits through a targeted savings strategy represents "
            f"a meaningful funding cost improvement opportunity.",
            ST["brokered_body"]))

    layout = Table(
        [[chart_tbl, Spacer(0.3 * inch, 1), [*narrative]]],
        colWidths=[5.1 * inch, 0.3 * inch, 4.0 * inch],
    )
    layout.setStyle(TableStyle([
        ("VALIGN",      (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING",  (0, 0), (-1, -1), 0),
    ]))
    story.append(layout)
    story.append(PageBreak())
    return story


def page_data(d, ST, narr={}):
    """Page 3 — What the data shows."""
    story = []
    story.append(Paragraph("WHAT THE DATA SHOWS", ST["section_label"]))
    story.append(Paragraph(
        f"{d['branches']} branches analyzed. "
        f"{d['at_risk']} require immediate strategic attention.",
        ST["page_headline"]))
    story.append(HRFlowable(width="100%", thickness=0.5,
                             color=TEAL, spaceAfter=8))
    # AI opening line for page 3
    _def_data_open = (
        f"Of {d['branches']} branches analyzed, {d['at_risk']} sit in "
        f"geographies where the bank is growing below the local market rate."
    )
    data_open   = narr.get("opening",   _def_data_open)
    data_ins    = narr.get("insight",   "")
    data_impl   = narr.get("implication", "")
    if data_open:
        story.append(Paragraph(data_open, ST["body"]))
    if data_ins:
        story.append(Paragraph(data_ins, ST["body_italic"]))
    story.append(Spacer(1, 6))

    # Zone tiles
    zone_data = [
        ("INVEST",   str(d["invest"]),   TEAL,     colors.HexColor("#EAF3DE"),
         "Branches with clear upside. Competitor is retreating or absent. Ready for campaign activation."),
        ("ANALYZE",  str(d["analyze"]),  colors.HexColor("#185FA5"), colors.HexColor("#E6F1FB"),
         "Contested markets. Growth is possible but requires sharper positioning and execution."),
        ("DEFEND",   str(d["defend"]),   AMBER,    colors.HexColor("#FFF3E0"),
         "Branches losing ground. Competitor pressure is active. Intervention is time-sensitive."),
        ("JUSTIFY",  str(d["justify"]),  RED_SOFT, colors.HexColor("#FCEBEB"),
         "Branches underperforming on all metrics. ROI of continued investment requires board review."),
    ]

    zone_rows = []
    for zone, count, fc, bg, desc in zone_data:
        tile = Table(
            [[Paragraph(f'<font color="#{fc.hexval()[2:]}" size="24"><b>{count}</b></font>',
                        ParagraphStyle("z", alignment=TA_CENTER, leading=28))],
             [Paragraph(f'<font color="#{fc.hexval()[2:]}" size="7"><b>{zone}</b></font>',
                        ParagraphStyle("zl", alignment=TA_CENTER, leading=9))]],
            colWidths=[0.7 * inch],
            rowHeights=[0.38 * inch, 0.20 * inch],
        )
        tile.setStyle(TableStyle([
            ("BACKGROUND",  (0, 0), (-1, -1), bg),
            ("TOPPADDING",  (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        desc_p = Paragraph(desc, ParagraphStyle("d",
            fontName="Helvetica", fontSize=9.5, textColor=NAVY,
            leading=14, alignment=TA_JUSTIFY))
        zone_rows.append([tile, desc_p])

    zone_tbl = Table(zone_rows,
        colWidths=[0.9 * inch, 8.3 * inch],
        spaceBefore=0)
    zone_tbl.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING",   (1, 0), (1, -1), 12),
        ("LINEBELOW",     (0, 0), (-1, -2), 0.3, GRAY2),
    ]))
    story.append(zone_tbl)
    story.append(Spacer(1, 0.15 * inch))

    # Volume at risk callout
    if d["vol_at_risk_M"] > 0:
        _def_callout = (
            f"${d['vol_at_risk_M']:.0f}M in deposits sit in branches "
            f"currently classified as Defend or Justify — locations growing below "
            f"the market rate in their own geographies."
            + (
                f" If the current trajectory continues, the gap implies "
                f"${d['proj_loss_M']:.0f}M in relative deposit underperformance "
                f"over the next 12 months."
                if d["proj_loss_M"] > 0 else ""
            )
        )
        callout_text = f"<b>{narr.get('implication', _def_callout) or _def_callout}</b>"
        callout_tbl = Table(
            [[Paragraph(callout_text, ST["callout"])]],
            colWidths=[9.2 * inch],
        )
        callout_tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), colors.HexColor("#FFF8EC")),
            ("LINEABOVE",     (0, 0), (-1, 0), 2.5, AMBER),
            ("TOPPADDING",    (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
            ("LEFTPADDING",   (0, 0), (-1, -1), 12),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 12),
        ]))
        story.append(callout_tbl)

    story.append(PageBreak())
    return story


def page_questions(d, ST, narr={}):
    """Page 4 — The strategic questions."""
    story = []
    story.append(Paragraph("THE STRATEGIC QUESTIONS", ST["section_label"]))
    story.append(Paragraph(
        "These are not marketing questions. They are balance sheet decisions.",
        ST["page_headline"]))
    story.append(HRFlowable(width="100%", thickness=0.5,
                             color=TEAL, spaceAfter=16))

    _def_q1 = (
        f"Of the {d['branches']} branches analyzed, {d['invest']} show clear competitive "
        f"upside with competitor retreat or absence. The other {d['analyze'] + d['defend'] + d['justify']} "
        f"require differentiated strategy — some need defense, some need triage. "
        f"Without branch-level scoring, capital and management attention are allocated by proximity, "
        f"not by opportunity."
    )
    _def_q2 = (
        f"{d['defend']} branches are actively losing ground to competitors operating in the same "
        f"geography. These are not slow-growth markets — they are contested markets where a "
        f"competitor is winning share. The question is not whether to act, but how urgently."
    )
    _def_q3 = (
        f"Cost of funds at {_sf(d['fin'].get('cost_of_funds_pct')):.2f}% reflects the blended "
        f"cost of the current deposit mix. "
        + (
            f"With {d['brok']['pct']}% of deposits sourced through brokers at a premium rate, "
            f"a targeted direct-acquisition strategy — savings account openers converting to CDs — "
            f"could meaningfully reduce this cost. "
            if d["brok"] else
            f"A targeted direct-acquisition strategy focused on savings accounts "
            f"and CD conversion could improve this cost structure over time. "
        )
        + "This is a funded account question, not a marketing question."
    )
    _def_close = (
        "Verlocity's branch-level intelligence platform provides decision-quality answers "
        "to all three questions — grounded in market data, not assumptions."
    )

    questions = [
        ("Which branches justify continued acquisition investment?",
         narr.get("q1_body", _def_q1)),
        ("Where should we be defending rather than growing?",
         narr.get("q2_body", _def_q2)),
        ("What is our true cost of deposits — and can it be reduced?",
         narr.get("q3_body", _def_q3)),
    ]

    for q, body in questions:
        story.append(KeepTogether([
            Paragraph(q, ST["question"]),
            Paragraph(body, ST["question_body"]),
        ]))

    story.append(Spacer(1, 0.1 * inch))
    story.append(Paragraph(
        narr.get("close", _def_close),
        ST["body_italic"]))

    story.append(PageBreak())
    return story


def page_ask(d, ST, narr={}):
    """Page 5 — The ask."""
    story = []
    story.append(Paragraph("THE RECOMMENDATION", ST["section_label"]))
    story.append(Paragraph(
        "Approve the strategic assessment. Understand your market before your competitors define it for you.",
        ST["page_headline"]))
    story.append(HRFlowable(width="100%", thickness=0.5,
                             color=TEAL, spaceAfter=14))

    _def_open = (
        f"Verlocity will deliver a comprehensive branch-level assessment of all "
        f"{d['branches']} locations in {d['bank_name']}'s network — where to invest, "
        f"where to defend, and where to exit. The output is a 12–15 page executive "
        f"report with a 90-day leadership agenda, delivered in 4–6 weeks."
    )
    _def_proof   = "One Verlocity client generated $660M in new deposits over two years using this approach."
    _def_urgency = "Competitors in your markets are not standing still — the cost of waiting grows each quarter."
    _def_close   = "The banks that define their deposit strategy now will set the terms for growth over the next three years."

    ask_open    = narr.get("opening",  _def_open)
    ask_proof   = narr.get("proof",    _def_proof)
    ask_urgency = narr.get("urgency",  _def_urgency)
    ask_close   = narr.get("close",    _def_close)

    story.append(Paragraph(ask_open, ST["body"]))
    story.append(Spacer(1, 0.06 * inch))

    if d["brok"]:
        story.append(Paragraph(
            f"The engagement will also quantify the funding cost opportunity "
            f"available through direct deposit acquisition versus the current "
            f"{d['brok']['pct']}% brokered position — and identify which branch "
            f"markets offer the highest-quality savings account acquisition potential.",
            ST["body"]))
        story.append(Spacer(1, 0.06 * inch))

    story.append(Paragraph(ask_proof,   ST["body_italic"]))
    story.append(Spacer(1, 0.04 * inch))
    story.append(Paragraph(ask_urgency, ST["body_italic"]))
    story.append(Spacer(1, 0.1 * inch))

    # Engagement table
    eng_data = [
        ["RECOMMENDED ACTION",  "Approve the Verlocity Strategic Assessment"],
        ["INVESTMENT",          "$12,000 – $15,000 (fixed fee)"],
        ["DELIVERABLE",         f"Branch scoring for all {d['branches']} locations + 90-day leadership agenda"],
        ["TIMELINE",            "4–6 weeks from engagement start"],
        ["NEXT STEP",           "Leadership session to confirm scope and begin Foundation phase"],
    ]
    eng_tbl = Table(eng_data, colWidths=[2.2 * inch, 7.0 * inch])
    eng_tbl.setStyle(TableStyle([
        ("FONTNAME",      (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (0, -1), 8),
        ("TEXTCOLOR",     (0, 0), (0, -1), GRAY3),
        ("FONTNAME",      (1, 0), (1, -1), "Helvetica-Bold"),
        ("FONTSIZE",      (1, 0), (1, -1), 11),
        ("TEXTCOLOR",     (1, 0), (1, -1), NAVY),
        ("TOPPADDING",    (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
        ("LEFTPADDING",   (0, 0), (-1, -1), 10),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
        ("LINEBELOW",     (0, 0), (-1, -2), 0.3, GRAY2),
        ("BACKGROUND",    (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR",     (0, 0), (-1, 0), WHITE),
        ("FONTNAME",      (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, 0), 9),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [GRAY1, WHITE]),
    ]))
    story.append(eng_tbl)
    story.append(Spacer(1, 0.25 * inch))

    story.append(Paragraph(ask_close, ST["body_italic"]))

    return story


# ══════════════════════════════════════════════════════════════
# MAIN BUILD
# ══════════════════════════════════════════════════════════════

def get_board_narratives(data):
    """
    Call Claude to generate CFO-quality prose for each page of the board brief.
    Returns dict with keys: situation, data, questions, ask
    Each value is a dict: {opening, insight, implication, close}
    """
    if not ANTH_KEY or not _anthropic:
        print("  ⚠  No ANTHROPIC_API_KEY — using placeholder narratives")
        empty = {"opening": "", "insight": "", "implication": "", "close": ""}
        return {k: empty for k in ["situation", "data", "questions", "ask"]}

    d = data

    def _sf(v, default=0):
        try:
            return float(v) if v is not None else default
        except (TypeError, ValueError):
            return default

    brok_str = (
        f"The bank has {d['brok']['pct']}% brokered deposit exposure "
        f"(${d['brok']['brokered_M']:.0f}M of ${d['brok']['total_M']:.0f}M total). "
        f"This is expensive, rate-sensitive capital — a funded account conversion "
        f"opportunity via savings→CD funnel targeting."
    ) if d.get("brok") else "Brokered deposits are not a material factor for this bank."

    ctx = f"""Bank: {d['bank_name']}
Total deposits: {d['deposits_str']} across {d['branches']} branches
Deposit YoY: {d['bank_yoy']:+.1f}% vs competitor avg in same markets: {d['comp_yoy']:+.1f}%
Gap vs market peers: {d['gap_str']}
Branch zones: Invest {d['invest']} | Analyze {d['analyze']} | Defend {d['defend']} | Justify {d['justify']}
At-risk branches (Defend+Justify): {d['at_risk']} holding ~${d['vol_at_risk_M']:.0f}M in deposits
Projected 12-month deposit underperformance if gap continues: ~${d['proj_loss_M']:.0f}M
ROA: {_sf(d['fin'].get('roa')):.2f}% | NIM: {_sf(d['fin'].get('nim')):.2f}% | Efficiency: {_sf(d['fin'].get('efficiency_ratio')):.2f}%
Cost of funds: {_sf(d['fin'].get('cost_of_funds_pct')):.2f}% | Net income YoY: {_sf(d['fin'].get('net_income_yoy_pct')):+.1f}%
{brok_str}"""

    system = """You are a senior financial strategist writing a confidential board brief for a bank's board of directors on behalf of Verlocity Princeton Partners Group.

TONE: Write like the bank's CFO wrote this — not an agency. Clinical, direct, numbers-first. No marketing language, no enthusiasm, no jargon. Boards are suspicious of anything that sounds like a sales document.

VERLOCITY CONTEXT:
- Verlocity is a Marketing Intelligence System — BMAP (branch scoring), Brandvention (brand strategy), Audiencefinder (targeted deposit campaigns), Clientdelight (loyalty/retention)
- The engagement being proposed: a strategic branch assessment ($12-15K fixed, 4-6 weeks) that tells the board exactly where to invest, defend, and exit across their network
- One Verlocity client generated $660M in new deposits over two years using this approach
- WinShare model: Verlocity's compensation is tied to results, not effort

CRITICAL RULES:
- Never name competitors. Say "a regional competitor" or "competitors in your markets"
- Never prescribe specific actions the bank can execute without Verlocity
- Frame everything as a question the board should be asking — not an answer
- The PDF purpose is to get board approval BEFORE the pitch meeting with Tom
- Each page should make NOT acting feel more expensive than acting
- Brokered deposits if present: frame as a balance sheet cost reduction opportunity, not a marketing campaign
- "Peer avg" = competitor deposit growth in the bank's own branch markets

Return ONLY valid JSON, no markdown:
{
  "situation": {
    "opening": "2 sentences. State the deposit gap fact cold. No framing, no softening.",
    "insight": "2 sentences. Why this gap matters to the balance sheet specifically — NIM, cost of funds, funding mix.",
    "implication": "1 sentence. What happens if this trajectory continues. Make it feel expensive.",
    "close": "1 sentence. The question the board should be asking that only Verlocity can answer."
  },
  "data": {
    "opening": "1 sentence. Frame the branch analysis as a balance sheet exercise, not a marketing exercise.",
    "insight": "2 sentences. What the zone distribution reveals about capital allocation and growth potential.",
    "implication": "1 sentence. The dollar amount at risk. Be specific.",
    "close": "1 sentence. What branch-level intelligence reveals that aggregate reporting cannot."
  },
  "questions": {
    "q1_body": "2 sentences answering why branch-level investment decisions matter for capital allocation.",
    "q2_body": "2 sentences on what active competitor pressure in 49 branches means for deposit stability.",
    "q3_body": "2 sentences on cost of funds and the funded account acquisition opportunity.",
    "close": "1 sentence. Position Verlocity as the only source of decision-quality answers to all three."
  },
  "ask": {
    "opening": "2 sentences. Frame the $12-15K assessment as the lowest-cost way to make a multi-million dollar decision correctly.",
    "proof": "1 sentence. Reference the $660M proof point without overselling.",
    "urgency": "1 sentence. Why timing matters — competitors are not standing still.",
    "close": "1 sentence. The closing line. Make not approving feel like the riskier choice."
  }
}"""

    print(f"  [board] Generating AI narratives...")
    client = _anthropic.Anthropic(api_key=ANTH_KEY)
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2000,
        system=system,
        messages=[{"role": "user", "content": f"Write board brief narratives for:\n\n{ctx}"}]
    )
    raw = msg.content[0].text.strip().replace("```json","").replace("```","").strip()
    try:
        narr = json.loads(raw)
        print(f"  [board] ✓ AI narratives generated")
        return narr
    except Exception as e:
        print(f"  [board] ⚠  JSON parse failed ({e}) — using placeholder narratives")
        empty = {"opening": "", "insight": "", "implication": "", "close": ""}
        return {k: empty for k in ["situation", "data", "questions", "ask"]}


def build_board_brief(data):
    """Build the board brief PDF and return bytes."""
    buf = io.BytesIO()
    ST  = make_styles()
    d   = data

    chrome = make_page_template(d["bank_name"], "CONFIDENTIAL — BOARD USE ONLY")

    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(letter),
        leftMargin=MARGIN + 0.18 * inch,  # account for navy stripe
        rightMargin=MARGIN,
        topMargin=MARGIN + 0.1 * inch,
        bottomMargin=0.6 * inch,
        title=f"Verlocity Board Brief — {d['bank_name']}",
        author="Verlocity Princeton Partners Group",
    )

    print(f"  [board] Fetching AI narratives...")
    narr = get_board_narratives(data)

    story = []
    story += page_cover(d, ST)
    story += page_situation(d, ST, narr.get("situation", {}))
    story += page_data(d, ST, narr.get("data", {}))
    story += page_questions(d, ST, narr.get("questions", {}))
    story += page_ask(d, ST, narr.get("ask", {}))

    doc.build(story, onFirstPage=chrome, onLaterPages=chrome)
    buf.seek(0)
    return buf


def generate_board_brief(ik, bank_name=None):
    """Top-level entry point — fetch data and build PDF. Returns BytesIO."""
    data = fetch_board_data(ik)
    if bank_name:
        data["bank_name"] = bank_name
    return build_board_brief(data)


# ══════════════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="BMAP Board Brief PDF Generator")
    parser.add_argument("--inst_key", required=True)
    parser.add_argument("--name",     default=None)
    args = parser.parse_args()

    print(f"\n{'='*55}")
    print(f"  Board Brief — {args.name or args.inst_key}")
    print(f"{'='*55}")

    buf  = generate_board_brief(args.inst_key, args.name)
    safe = "".join(c if c.isalnum() or c in " _-" else "_"
                   for c in (args.name or args.inst_key)).strip()
    date = datetime.now().strftime("%Y%m%d")
    out  = OUT_DIR / f"Board_Brief_{safe}_{date}.pdf"
    out.write_bytes(buf.read())
    print(f"\n  ✓  Saved: {out}\n")
