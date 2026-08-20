"""
bmap_branch_preview.py — Single-Branch Deep Dive Preview

Purpose: a LIVE PITCH-MEETING tool, distinct from both:
  - bmap_snapshot.py    (automated cold/warm prospecting outreach)
  - bmap_assessment_doc.py (the paid $10K full-network deliverable)

This shows a prospect ONE branch's full deep-dive treatment, at the exact
same depth and fidelity as the paid Assessment, to demonstrate what the
full network engagement produces. It is NOT a simplified mockup: it
literally reuses render_branch_deep_dive() from bmap_assessment_doc.py,
the same function that renders every branch in the real paid deliverable.
If that function's output quality changes (better or worse), this preview
changes with it automatically -- there is no separate, divergent code path
to keep in sync.

Usage:
    python bmap_branch_preview.py --inst_key bank_463735 --name "Hancock Whitney Bank"
    python bmap_branch_preview.py --inst_key bank_463735 --branch "Gulfport Main Branch"

If --branch is omitted, the most compelling branch is auto-selected:
flagship risk first (if one exists), otherwise the top opportunity-score
branch -- so a sales rep never has to know the network well enough to
pick a branch themselves.
"""

import os
import sys
import argparse
from pathlib import Path
from datetime import datetime

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Reuses bmap_assessment_doc.py's actual rendering code, styling constants,
# and data-fetch functions -- co-located in the same repo/deploy directory.
import bmap_assessment_doc as bad


def pick_preview_branch(summary, branches, branch_name=None):
    """Selects which branch to preview. Explicit --branch always wins (exact
    or case-insensitive substring match on namebr). Otherwise auto-picks the
    most compelling story available: the flagship-risk branch if one exists
    (a bank's biggest branch in real trouble is the most attention-grabbing
    opener in a pitch), falling back to the top opportunity-score branch."""
    if branch_name:
        needle = branch_name.strip().lower()
        exact = [b for b in branches if (b.get("namebr") or "").strip().lower() == needle]
        if exact:
            return exact[0]
        partial = [b for b in branches if needle in (b.get("namebr") or "").lower()]
        if len(partial) == 1:
            return partial[0]
        if len(partial) > 1:
            names = ", ".join(b.get("namebr") for b in partial)
            raise ValueError(f"'{branch_name}' matches multiple branches: {names}. Be more specific.")
        raise ValueError(f"No branch matching '{branch_name}' found in this network.")

    if summary.get("flagship_risk"):
        fr = summary["flagship_risk"]
        match = [b for b in branches if b.get("uninumbr") == fr.get("uninumbr")]
        if match:
            return match[0]
    if summary.get("top5"):
        return summary["top5"][0]
    return branches[0] if branches else None


def get_single_branch_narrative(bank_name, branch, strat, play):
    """Lightweight narrative call scoped to ONE branch -- deliberately NOT
    the full get_narratives() network-wide schema (exec_headline, priority_focus,
    next_12_months, etc. would be wasted tokens/latency for a single-branch
    preview meant to be fast enough to run live in a pitch meeting). Returns
    the same {"branch_verdicts": ..., "branch_plays": ..., "branch_audiences": ...}
    shape render_branch_deep_dive() already expects, so no changes needed there."""
    branch_label = f"{branch.get('namebr')} ({branch.get('citybr')}, {branch.get('stalpbr')})"
    empty = {"branch_verdicts": {}, "branch_plays": {}, "branch_audiences": {}}

    if not bad.ANTH_KEY or not bad.anthropic:
        return empty

    top_comp = strat.get("top_competitor") if strat else None
    comp_str = (f"nearest named competitor {top_comp.get('bank_name')} "
                f"{bad._sf(top_comp.get('distance_miles')):.2f}mi away with "
                f"${bad._sf(top_comp.get('deposits'))/1e6:.0f}M deposits"
                if top_comp else "no named competitor within the adaptive radius")

    ctx = (
        f"Bank: {bank_name}\n"
        f"Branch: {branch_label}\n"
        f"Score {bad._sf(branch.get('opportunity_score')):.0f}/100, "
        f"zone {branch.get('opportunity_zone')}, "
        f"${bad._sf(branch.get('latest_dep'))/1e6:.0f}M deposits, "
        f"{bad._sf(branch.get('yoy_deposits'))*100:+.1f}% YoY, {comp_str}.\n"
        f"Household income ${bad._sf(branch.get('household_income')):.0f} "
        f"({bad._sf(branch.get('yoy_income_growth'))*100:+.1f}% YoY), "
        f"population YoY {bad._sf(branch.get('yoy_pop_growth'))*100:+.1f}%, "
        f"home value YoY {bad._sf(branch.get('zhvi_yoy_pct')):+.1f}%.\n"
        f"Assigned play: {play or 'n/a'}."
    )

    system = """You are writing a ONE-BRANCH preview of the Verlocity BMAP Assessment,
shown live in a sales pitch to demonstrate analytical depth. Same standards as the full
paid Assessment: confident, commercial, decisive, grounded in the specific numbers given.
Every claim needs a number. Return ONLY valid JSON, no markdown fences:
{
  "branch_verdicts": {"Branch Name (City, ST)": "3-4 sentences synthesizing score, zone, the named competitive threat (or its absence), and deposit trajectory into a clear verdict -- the 'why' behind the assigned play."},
  "branch_plays": {"Branch Name (City, ST)": {"resource_posture": "One sentence, grounded in this branch's specific numbers -- not a generic play-name restatement. If no named competitor exists within the adaptive radius, do NOT write language implying one does.", "media_brief": "One to two sentences, naming the actual target audience and product implied by this branch's specific data."}},
  "branch_audiences": {"Branch Name (City, ST)": "2-3 sentences using ONLY the household income, income YoY, population YoY, and home value YoY figures given. Frame through Verlocity's AudienceFinder segments (High-Quality Local Prospect, Regression-Scored Lookalike, Competitive Conquesting, Warm Retargeting) where the signal supports it. Never invent a named persona."}
}"""

    try:
        client = bad.anthropic.Anthropic(api_key=bad.ANTH_KEY)
        msg = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1500,
            thinking={"type": "disabled"},
            system=system,
            messages=[{"role": "user", "content": ctx}],
        )
        raw = msg.content[0].text.strip().replace("```json", "").replace("```", "").strip()
        import json
        return json.loads(raw)
    except Exception as e:
        print(f"  ⚠ Single-branch narrative failed ({type(e).__name__}: {e}) — using deterministic fallback")
        return empty


def build_branch_preview_doc(bank_name, branch, strat, play, entry, capped_yoy, narr,
                              total_branch_count, geo_by_uid, tmpdir="."):
    """Short standalone document: intro framing -> the one branch's full deep
    dive (via bad.render_branch_deep_dive, identical to the paid Assessment)
    -> closing CTA. Uses the same brand styling as bmap_assessment_doc.py."""
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.59)
    section.page_height = Cm(27.94)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)

    # ── Cover / intro ──
    p0 = doc.add_paragraph()
    p0.paragraph_format.space_before = Pt(50)
    logo_path = str(Path(__file__).parent / "verlocity_logo.jpg")
    if os.path.exists(logo_path):
        p0.add_run().add_picture(logo_path, width=Inches(2.0))
    else:
        r0 = p0.add_run("VERLOCITY")
        r0.bold = True
        r0.font.size = Pt(14)
        r0.font.color.rgb = bad.TEAL
        r0.font.name = bad.FONT_HEAD

    p1 = doc.add_paragraph()
    p1.paragraph_format.space_before = Pt(18)
    r1 = p1.add_run(bank_name)
    r1.bold = True
    r1.font.size = Pt(28)
    r1.font.color.rgb = bad.NAVY
    r1.font.name = bad.FONT_HEAD

    p2 = doc.add_paragraph()
    p2.paragraph_format.space_after = Pt(6)
    r2 = p2.add_run("BMAP Assessment — Branch Deep Dive Preview")
    r2.font.size = Pt(14)
    r2.font.color.rgb = bad.GRAY3
    r2.font.name = bad.FONT_HEAD

    p3 = doc.add_paragraph()
    p3.paragraph_format.space_after = Pt(30)
    r3 = p3.add_run(datetime.now().strftime("%B %Y"))
    r3.font.size = Pt(11)
    r3.font.color.rgb = bad.GRAY3
    r3.font.name = bad.FONT_HEAD

    p_intro = doc.add_paragraph()
    p_intro.paragraph_format.space_after = Pt(20)
    r_intro = p_intro.add_run(
        f"This is one branch, shown at the exact depth and analytical rigor of the full "
        f"BMAP Assessment — the same competitive geocoding, adaptive-radius modeling, "
        f"and capture-dollar sizing your team would receive for every priority branch "
        f"in {bank_name}'s {total_branch_count}-branch network."
    )
    r_intro.italic = True
    r_intro.font.size = Pt(11)
    r_intro.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    r_intro.font.name = bad.FONT_HEAD

    doc.add_page_break()

    # ── The one branch, full depth — identical code path to the paid Assessment ──
    bad.render_branch_deep_dive(
        doc, branch, strat, play, entry, capped_yoy,
        narr.get("branch_verdicts") or {}, narr.get("branch_plays") or {},
        narr.get("branch_audiences") or {}, geo_by_uid,
        tmpdir, heading_space_before=0,
    )

    # ── Closing CTA ──
    doc.add_page_break()
    p_cta_h = doc.add_paragraph()
    p_cta_h.paragraph_format.space_before = Pt(20)
    r_cta_h = p_cta_h.add_run("This Is One Branch.")
    r_cta_h.bold = True
    r_cta_h.font.size = Pt(20)
    r_cta_h.font.color.rgb = bad.NAVY
    r_cta_h.font.name = bad.FONT_HEAD

    cta_box = doc.add_table(rows=1, cols=1)
    cell = cta_box.rows[0].cells[0]
    bad._set_cell_shading(cell, "083D5F")
    cell.paragraphs[0].text = ""
    r_cta = cell.paragraphs[0].add_run(
        f"{bank_name}'s network has {total_branch_count} branches. The full BMAP Assessment "
        f"delivers this exact depth — verdict, competitive radius map, capture-dollar modeling, "
        f"audience signal, assigned play — for every priority branch, plus network-wide "
        f"executive synthesis, live market intelligence, and financial benchmarking."
    )
    r_cta.font.size = Pt(12)
    r_cta.font.bold = True
    r_cta.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    r_cta.font.name = bad.FONT_HEAD
    cell.paragraphs[0].paragraph_format.space_before = Pt(10)
    cell.paragraphs[0].paragraph_format.space_after = Pt(10)

    return doc


def generate_preview(ik, name_hint=None, branch_name=None, tmpdir="."):
    """Core logic, returns (doc, bank_name, branch_label, total_branch_count)
    or raises ValueError with a clear message. Separated from run() so both
    the CLI entry point (saves to disk) and the Hub's Quick Export endpoint
    (streams from memory, same pattern as /generate-assessment) can share
    it without duplicating the fetch/pick/narrate/build sequence."""
    print(f"[branch-preview] fetching network data for {ik}...")
    d = bad.fetch_full_network_data(ik)
    if not d["branches"]:
        raise ValueError(f"No branch data found for inst_key='{ik}'. This bank may not be "
                          f"ingested into BMAP yet, or the inst_key is incorrect.")

    bank_name = name_hint or (d["branches"][0].get("namefull") if d["branches"] else None) or ik
    branches = d["branches"]
    branch_strategy = d.get("branch_strategy") or []
    capped_yoy = d.get("capped_yoy") or {}
    geo_by_uid = {g["uninumbr"]: g for g in (d.get("branches_geo") or []) if g.get("uninumbr") is not None}

    summary = bad.summarize_network(d)
    dives, _ = bad.build_branch_deep_dives(branches, branch_strategy)

    target = pick_preview_branch(summary, branches, branch_name)
    if not target:
        raise ValueError("Could not select a branch to preview — network has no branches.")

    entry = next((e for e in dives if e["branch"].get("uninumbr") == target.get("uninumbr")), None)
    if not entry:
        # Target branch wasn't in the curated top-15 (large network) --
        # build its entry directly so the preview isn't limited to only
        # branches the full Assessment happened to curate.
        strategy_by_name = {(r["namebr"], r["citybr"], r["stalpbr"]): r for r in branch_strategy}
        key = (target.get("namebr"), target.get("citybr"), target.get("stalpbr"))
        strat = strategy_by_name.get(key)
        play = bad.get_play(target.get("opportunity_zone"), target.get("matrix_quadrant"))
        top_comp = strat.get("top_competitor") if strat else None
        capture_pool = bad._sf(top_comp.get("deposits")) if top_comp else 0.0
        entry = {"branch": target, "strategy": strat, "play": play, "capture_pool": capture_pool}

    branch, strat, play = entry["branch"], entry["strategy"], entry["play"]
    branch_label = f"{branch.get('namebr')} ({branch.get('citybr')}, {branch.get('stalpbr')})"
    print(f"[branch-preview] previewing {branch_label}")

    narr = get_single_branch_narrative(bank_name, branch, strat, play)
    doc = build_branch_preview_doc(bank_name, branch, strat, play, entry, capped_yoy, narr,
                                    len(branches), geo_by_uid, tmpdir=tmpdir)
    return doc, bank_name, branch_label, len(branches)


def run(ik, name_hint=None, branch_name=None):
    import tempfile
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            doc, bank_name, branch_label, total = generate_preview(ik, name_hint, branch_name, tmpdir)
        except ValueError as e:
            print(f"[branch-preview] ✗ {e}")
            return None

        safe = "".join(c if c.isalnum() or c in " _-" else "_" for c in bank_name).strip()
        date = datetime.now().strftime("%Y%m%d")
        branch_safe = "".join(c if c.isalnum() or c in " _-" else "_"
                               for c in branch_label.split(" (")[0]).strip()
        filename = f"BMAP_Preview_{safe}_{branch_safe}_{date}.docx"
        out_path = os.path.join(bad.OUT_DIR if hasattr(bad, "OUT_DIR") else ".", filename)
        doc.save(out_path)

    print(f"[branch-preview] ✓ {out_path}")
    return out_path


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate a single-branch BMAP Assessment preview")
    parser.add_argument("--inst_key", required=True, help="e.g. bank_463735")
    parser.add_argument("--name", default=None, help="Bank display name (optional)")
    parser.add_argument("--branch", default=None,
                         help="Branch name to preview (optional -- auto-selects the most "
                              "compelling branch if omitted: flagship risk, else top opportunity score)")
    args = parser.parse_args()
    result = run(args.inst_key, args.name, args.branch)
    sys.exit(0 if result else 1)
