"""Pure engagement logic: field validation, calculated dates and fees,
assertions, and assembly of the SOW section model.

Nothing here talks to a database, an LLM or the filesystem. Clause text,
module text and prices arrive as a `library` dict (loaded from the private
public.engagement_* tables by store.get_library), so this module can be
tested with a synthetic library.

library = {
  "version": int,
  "clauses":  {section: [{"seq", "style", "body", "when_modules"}, ...]},
  "modules":  {key: {"name", "fee_label", "price", "from_month", "to_month",
                     "depends_on", "activities", "tail_activities",
                     "deliverable", "sort"}},
  "module_text": {key: [{"seq", "style", "body"}, ...]},
}
"""

import calendar
import datetime as dt
import hashlib
import json
import re

INSTALLMENT_ROUND = 1000  # derived from the reference SOW: 4 x 20,000 + 21,750 on 101,750
MIN_TERM, MAX_TERM = 3, 12
ADVISORY = "advisory"
CAMPAIGN_STEPS = ("steps14", "step5", "step6")
STEP_NUMBER = {"steps14": 4, "step5": 5, "step6": 6}
TERM_WORDS = {1: "one", 2: "two", 3: "three", 4: "four", 5: "five", 6: "six", 7: "seven",
              8: "eight", 9: "nine", 10: "ten", 11: "eleven", 12: "twelve"}
REQUIRED_SECTIONS = ["delivery", "governance", "responsibilities", "fees_intro", "fees_intro_campaign",
                     "advisory_note", "advisory_note_campaign", "third_party_pass", "third_party_direct",
                     "assumptions", "outofscope", "scope_changes", "acceptance", "term", "status",
                     "scope_intro", "workplan_intro", "objectives_std"]
BANNED_IN_NARRATIVE = ["guarantee", "guaranteed", "guarantees", "promise", "promises", "assure", "assured"]
START_RE = re.compile(r"^(20\d\d)-(0[1-9]|1[0-2])-01$")
DATE_RE = re.compile(r"^(20\d\d)-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$")
PLACEHOLDER_RE = re.compile(r"\{[a-z_]+\}")


def cap(s):
    return s[:1].upper() + s[1:]


def long_date(d):
    return f"{calendar.month_name[d.month]} {d.day}, {d.year}"


def month_year(d):
    return f"{calendar.month_name[d.month]} {d.year}"


def add_months(d, n):
    """First day of the month n months after d (d is a first-of-month date)."""
    idx = d.year * 12 + (d.month - 1) + n
    return dt.date(idx // 12, idx % 12 + 1, 1)


def term_end(start, term):
    return add_months(start, term) - dt.timedelta(days=1)


def narrative_key(fields):
    basis = [fields.get("client_name"), fields.get("client_short"),
             fields.get("context") or [], fields.get("objectives") or []]
    return hashlib.sha1(json.dumps(basis, sort_keys=True).encode()).hexdigest()[:12]


# ── narrative guardrails ───────────────────────────────────────────────
def check_narrative(parts):
    """Returns a list of failure strings for LLM-drafted text. Empty list = ok.
    The drafting step may not introduce numbers, guarantees or placeholders."""
    text = " ".join(p for p in parts if isinstance(p, str))
    fails = []
    if re.search(r"\d", text):
        fails.append("contains a number or digit")
    low = text.lower()
    for w in BANNED_IN_NARRATIVE:
        if re.search(r"\b" + w + r"\b", low):
            fails.append(f"contains the banned word '{w}'")
    if "{" in text or "}" in text:
        fails.append("contains a placeholder brace")
    return fails


# ── fields ─────────────────────────────────────────────────────────────
def choosable_modules(library):
    return [k for k, m in sorted(library["modules"].items(), key=lambda kv: kv[1]["sort"]) if k != ADVISORY]


def order_modules(keys, library):
    chosen = {k for k in keys if k in library["modules"] and k != ADVISORY}
    ordered = [k for k in choosable_modules(library) if k in chosen]
    if ADVISORY in library["modules"]:
        ordered.append(ADVISORY)
    return ordered


def clean_fields(raw, library):
    """Normalize operator/LLM-provided fields. Returns (fields, errors).
    errors is a list of strings describing problems that block generation;
    fields always holds the best-effort normalized values."""
    raw = raw or {}
    errors = []
    f = {}
    f["client_name"] = str(raw.get("client_name") or "").strip()[:120]
    f["client_short"] = (str(raw.get("client_short") or "").strip() or f["client_name"])[:60]
    f["sponsor"] = str(raw.get("sponsor") or "").strip()[:80]
    mods = [m for m in (raw.get("modules") or []) if isinstance(m, str)]
    unknown = [m for m in mods if m not in library["modules"]]
    f["modules"] = order_modules(mods, library)
    if unknown:
        errors.append("Unknown services: " + ", ".join(sorted(set(unknown))))
    start = raw.get("start")
    f["start"] = start if isinstance(start, str) and START_RE.match(start) else None
    if start and not f["start"]:
        errors.append("Start must be the first day of a month (YYYY-MM-01)")
    term = raw.get("term")
    f["term"] = term if isinstance(term, int) and not isinstance(term, bool) and MIN_TERM <= term <= MAX_TERM else None
    if term is not None and f["term"] is None:
        errors.append(f"Term must be a whole number of months from {MIN_TERM} to {MAX_TERM}")
    tp = raw.get("third_party")
    f["third_party"] = tp if tp in ("pass", "direct") else None
    f["context"] = _strings(raw.get("context"))
    f["objectives"] = _strings(raw.get("objectives"))
    f["scope_overrides"] = _strings(raw.get("scope_overrides"), maxn=6)
    dd = raw.get("draft_date")
    f["draft_date"] = dd if isinstance(dd, str) and DATE_RE.match(dd) else dt.date.today().isoformat()
    sg = raw.get("signers") if isinstance(raw.get("signers"), dict) else {}
    f["signers"] = {side: {"name": str((sg.get(side) or {}).get("name") or "")[:80],
                           "title": str((sg.get(side) or {}).get("title") or "")[:80]}
                    for side in ("client", "verlocity")}
    nar = raw.get("narrative")
    f["narrative"] = nar if isinstance(nar, dict) and isinstance(nar.get("purpose"), list) else None
    # dependency check (errors, not silent drops: the operator must see it)
    for k in f["modules"]:
        dep = library["modules"][k].get("depends_on")
        if dep and dep not in f["modules"]:
            errors.append(f"{library['modules'][k]['name']} requires {library['modules'][dep]['name']}")
    return f, errors


def _strings(v, maxn=8):
    out = []
    for x in (v or []):
        if not isinstance(x, str):
            continue
        x = x.strip()
        if x and len(x) <= 160 and x.lower() not in [y.lower() for y in out] and len(out) < maxn:
            out.append(x)
    return out


def drop_unmet_dependencies(modules, library):
    """For conversational use: remove modules whose prerequisite is absent.
    Returns (kept, removed_names)."""
    kept = list(modules)
    removed = []
    changed = True
    while changed:
        changed = False
        for k in list(kept):
            dep = library["modules"][k].get("depends_on")
            if dep and dep not in kept:
                kept.remove(k)
                removed.append(library["modules"][k]["name"])
                changed = True
    return kept, removed


# ── fees ───────────────────────────────────────────────────────────────
def compute_fees(modules, term, library):
    comps = []
    for k in modules:
        m = library["modules"][k]
        if k == ADVISORY:
            continue
        comps.append({"key": k, "label": m["fee_label"], "amount": int(m["price"])})
    total = sum(c["amount"] for c in comps)
    base = (total // term // INSTALLMENT_ROUND) * INSTALLMENT_ROUND if term else 0
    last = total - base * (term - 1) if term else total
    installments = [base] * (term - 1) + [last] if term else []
    return {"components": comps, "total": total, "base": base, "last": last, "installments": installments}


def installment_rows(fees, term):
    base, last = fees["base"], fees["last"]
    money = lambda n: f"${n:,}"
    if term == 1 or base == last:
        rows = [[f"Months 1–{term}", f"{money(base)} per month"]]
    elif term == 2:
        rows = [["Month 1", money(base)], ["Month 2", money(last)]]
    else:
        pre = f"Months 1–{term - 1}" if term > 2 else "Month 1"
        rows = [[pre, f"{money(base)} per month"], [f"Month {term}", money(last)]]
    rows.append(["Total", money(fees["total"])])
    return rows


# ── workplan ───────────────────────────────────────────────────────────
def build_workplan(modules, term, start, library):
    rows = []
    for i in range(1, term + 1):
        mstart = add_months(start, i - 1)
        active, acts, dels = [], [], []
        for k in modules:
            m = library["modules"][k]
            if k == ADVISORY:
                acts.append(m.get("activities") or "")
                continue
            if m["from_month"] <= i <= m["to_month"]:
                active.append(m["name"].split(" (")[0])
                acts.append(m.get("activities") or "")
            elif i > m["to_month"] and m.get("tail_activities"):
                acts.append(m["tail_activities"])
            if m["to_month"] == i and m.get("deliverable"):
                dels.append(m["deliverable"])
        if i == 1:
            acts.insert(0, "Kickoff; integrated workplan; data and document request; governance")
            dels.insert(0, "Confirmed workplan and governance; data-gap and dependency log")
        if i == term:
            dels.append("Executive closeout and proposed scope for any subsequent engagement")
        if i == 1:
            emphasis = "Mobilization and foundation"
        elif i == term:
            emphasis = "Results, decisions, and transition"
        else:
            emphasis = "; ".join(active) if active else "Advisory and learning"
        rows.append([f"Month {i}\n{month_year(mstart)}", emphasis,
                     "; ".join(a for a in acts if a), "; ".join(dels)])
    return rows


# ── params and text assembly ───────────────────────────────────────────
def workstream_count(mods):
    n = 1  # advisory is always a workstream
    if "assess" in mods:
        n += 1
    if any(k in mods for k in CAMPAIGN_STEPS):
        n += 1
    return n


def build_params(fields, library, fees):
    start = dt.date.fromisoformat(fields["start"])
    term = fields["term"]
    end = term_end(start, term)
    mods = fields["modules"]
    campaign = [k for k in mods if k in CAMPAIGN_STEPS]
    top = max((STEP_NUMBER[k] for k in campaign), default=0)
    scope_parts = []
    if "assess" in mods:
        scope_parts.append("the Assessment")
    if campaign:
        scope_parts.append(f"Campaign Steps 1–{top}")
    return {
        "client": fields["client_short"],
        "client_name": fields["client_name"],
        "sponsor_phrase": fields["sponsor"] or "the executive sponsor",
        "heavy_months": f"{calendar.month_name[start.month]} and {calendar.month_name[add_months(start, 1).month]}",
        "term_words": TERM_WORDS[term],
        "start_month": calendar.month_name[start.month],
        "start_month_year": month_year(start),
        "workstream_count_words": TERM_WORDS.get(workstream_count(mods), str(workstream_count(mods))),
        "end_month_year": month_year(end),
        "start_long": long_date(start),
        "end_long": long_date(end),
        "fee_scope": " and ".join(scope_parts) or "the selected services",
        "campaign_steps": f"Steps 1–{top}" if top else "",
        "total": f"${fees['total']:,}",
    }


def fmt(text, params):
    def rep(m):
        k = m.group(0)[1:-1]
        return str(params[k]) if k in params else m.group(0)
    return PLACEHOLDER_RE.sub(rep, text)


def clauses_for(library, section, selected, params):
    """Ordered blocks for a clause section, filtered by when_modules."""
    out = []
    for c in sorted(library["clauses"].get(section, []), key=lambda c: c["seq"]):
        need = set(c.get("when_modules") or [])
        if need <= set(selected):
            out.append({"t": c["style"], "x": fmt(c["body"], params)})
    return out


def _included_text(fields, params):
    mods = fields["modules"]
    parts = []
    if "assess" in mods:
        parts.append("Deposit Franchise Assessment")
    campaign = [k for k in mods if k in CAMPAIGN_STEPS]
    if campaign:
        top = max(STEP_NUMBER[k] for k in campaign)
        label = f"Campaign Steps 1–{top}"
        if "step6" in mods:
            label += ", including creative development and production"
        parts.append(label)
        if "step6" in mods:
            parts.append("campaign management")
    parts.append("and ongoing strategic advisory described in this SOW")
    return "; ".join(parts)


def build_sections(fields, library, fees, params, narrative_ok):
    mods = fields["modules"]
    sel = set(mods)
    term = fields["term"]
    campaign = [k for k in mods if k in CAMPAIGN_STEPS]
    suffix = "_campaign" if campaign else ""
    money = lambda n: f"${n:,}"
    sections = []

    # 1 purpose, 2 objectives (tailored by Claude; fixed std objectives from the library)
    nar = fields.get("narrative") if narrative_ok else None
    if nar:
        purpose = [{"t": "p", "x": p} for p in nar["purpose"]]
    else:
        purpose = [{"t": "p", "x": "[Purpose to be drafted from the discovery conversation.]", "flag": True}]
    sections.append({"title": "1. Engagement Purpose", "blocks": purpose})

    std = clauses_for(library, "objectives_std", sel, params)
    tailored = [{"t": "bullet", "x": o} for o in (nar["objectives"] if nar else [])]
    if len(std) > 1:
        objs = std[:1] + tailored + std[1:]
    else:
        objs = std + tailored
    if not nar and not std:
        objs = [{"t": "bullet", "x": "[Objectives to be drafted.]", "flag": True}]
    sections.append({"title": "2. Engagement Objectives", "blocks": objs})

    # 3 scope: workstreams from modules
    ws = []
    if "assess" in sel:
        ws.append(("Deposit Franchise Assessment", ["assess"]))
    if campaign:
        ws.append(("Consumer Deposit Campaign", campaign))
    ws.append(("Ongoing Strategic Advisory Partnership", [ADVISORY]))
    pillars = 5 if "assess" in sel else 0
    title3 = f"3. Integrated Scope of Services ({len(ws)} Workstreams" + (f", {pillars} Pillars)" if pillars else ")")
    blocks3 = clauses_for(library, "scope_intro", sel, params)
    for i, (name, keys) in enumerate(ws):
        blocks3.append({"t": "h3", "x": f"3.{i + 1} Workstream {chr(65 + i)}: {name}"})
        for k in keys:
            for b in sorted(library["module_text"].get(k, []), key=lambda b: b["seq"]):
                blocks3.append({"t": b["style"], "x": fmt(b["body"], params)})
    sections.append({"title": title3, "blocks": blocks3})

    sections.append({"title": "4. Delivery Approach", "blocks": clauses_for(library, "delivery", sel, params)})

    wp_blocks = clauses_for(library, "workplan_intro", sel, params)
    wp_blocks.append({"t": "table", "header": ["Month", "Primary Emphasis", "Principal Activities",
                                               "Expected Deliverables / Decisions"],
                      "rows": build_workplan(mods, term, dt.date.fromisoformat(fields["start"]), library),
                      "widths": [1.0, 1.4, 2.6, 2.6]})
    sections.append({"title": f"5. {cap(params['term_words'])}-Month Workplan and Deliverables", "blocks": wp_blocks})

    sections.append({"title": "6. Governance and Working Cadence", "blocks": clauses_for(library, "governance", sel, params)})
    sections.append({"title": f"7. {params['client']} Responsibilities", "blocks": clauses_for(library, "responsibilities", sel, params)})

    fb = clauses_for(library, "fees_intro" + suffix, sel, params)
    rows = [[c["label"], money(c["amount"])] for c in fees["components"]]
    rows.append(["Total professional fees", money(fees["total"])])
    fb.append({"t": "table", "header": ["Professional Fee Component", "Amount"], "rows": rows, "widths": [5.0, 2.0], "bold_last": True})
    fb.append({"t": "table", "header": ["Installment Schedule", "Amount"], "rows": installment_rows(fees, term),
               "widths": [5.0, 2.0], "bold_last": True})
    fb += clauses_for(library, "advisory_note" + suffix, sel, params)
    fb += clauses_for(library, "third_party_" + (fields["third_party"] or "pass"), sel, params)
    sections.append({"title": "8. Professional Fees and Third-Party Costs", "blocks": fb})

    sections.append({"title": "9. Assumptions and Dependencies", "blocks": clauses_for(library, "assumptions", sel, params)})
    oos = clauses_for(library, "outofscope", sel, params)
    oos += [{"t": "bullet", "x": o} for o in fields.get("scope_overrides") or []]
    sections.append({"title": "10. Out-of-Scope Services", "blocks": oos})
    sections.append({"title": "11. Scope Changes", "blocks": clauses_for(library, "scope_changes", sel, params)})
    sections.append({"title": "12. Review and Acceptance of Deliverables", "blocks": clauses_for(library, "acceptance", sel, params)})
    sections.append({"title": "13. Term, Termination, and Subsequent Work", "blocks": clauses_for(library, "term", sel, params)})
    sections.append({"title": "14. Authorization", "blocks": [{"t": "sign", "client": fields["client_name"].upper(),
                                                              "signers": fields["signers"]}]})
    return sections


def summary_rows(fields, library, fees, params):
    term, n = fields["term"], fields["term"]
    money = lambda x: f"${x:,}"
    if fees["base"] == fees["last"]:
        pay = f"{money(fees['total'])} total, paid over {params['term_words']} months in equal installments of {money(fees['base'])}"
    else:
        pay = (f"{money(fees['total'])} total, paid over {params['term_words']} months: "
               f"{TERM_WORDS.get(n - 1, str(n - 1))} installments of {money(fees['base'])} and a final installment of {money(fees['last'])}")
    third = ("Separately authorized pass-through expenses; billed separately" if fields["third_party"] != "direct"
             else f"Contracted and paid directly by {params['client']}; not billed by Verlocity")
    status = " ".join(b["x"] for b in clauses_for(library, "status", set(fields["modules"]), params)) or "Draft"
    return [
        ["Client", f'{fields["client_name"]} ("{params["client"]}")'],
        ["Service Provider", 'Verlocity, LLC ("Verlocity")'],
        ["Term", f"{cap(params['term_words'])} months; anticipated {params['start_long']} through {params['end_long']}"],
        ["Professional Fee", pay],
        ["Included", _included_text(fields, params)],
        ["Third-Party Costs", third],
        ["Draft Date", long_date(dt.date.fromisoformat(fields["draft_date"]))],
        ["Status", status],
    ]


# ── assertions ─────────────────────────────────────────────────────────
def _a(level, text, aid):
    return {"id": aid, "level": level, "text": text}


def missing_inputs(fields):
    miss = []
    if not fields["client_name"]:
        miss.append("client name")
    if not [m for m in fields["modules"] if m != ADVISORY]:
        miss.append("services")
    if not fields["start"]:
        miss.append("start date")
    if not fields["term"]:
        miss.append("term")
    if not fields["third_party"]:
        miss.append("third-party cost arrangement")
    return miss


def build_plan(raw_fields, library):
    """Main entry. Returns a plan dict; plan['ready'] is False when inputs are
    missing or any assertion fails (generation is blocked)."""
    fields, errors = clean_fields(raw_fields, library)
    assertions = []
    for e in errors:
        assertions.append(_a("fail", e, "input"))
    missing = missing_inputs(fields)
    if missing:
        return {"ready": False, "fields": fields, "missing": missing, "assertions": assertions,
                "library_version": library.get("version", 1)}

    absent = [s for s in REQUIRED_SECTIONS if s not in library["clauses"]]
    assertions.append(_a("fail" if absent else "pass",
                         "Clause library is complete" if not absent else "Clause library is missing: " + ", ".join(absent), "library"))
    if absent or errors:
        return {"ready": False, "fields": fields, "missing": [], "assertions": assertions,
                "library_version": library.get("version", 1)}

    term = fields["term"]
    start = dt.date.fromisoformat(fields["start"])
    fees = compute_fees(fields["modules"], term, library)
    params = build_params(fields, library, fees)

    # fees
    comp_sum = sum(c["amount"] for c in fees["components"])
    assertions.append(_a("pass" if comp_sum == fees["total"] else "fail",
                         f"Fee components add up to the total (${comp_sum:,} of ${fees['total']:,})", "fee_components"))
    ins_sum = sum(fees["installments"])
    assertions.append(_a("pass" if ins_sum == fees["total"] and len(fees["installments"]) == term and all(x >= 0 for x in fees["installments"]) else "fail",
                         f"Installments add up to the total (${ins_sum:,} of ${fees['total']:,}) across {term} months", "installments"))
    priced = [library["modules"][k]["name"] for k in fields["modules"] if k != ADVISORY and library["modules"][k]["price"] <= 0]
    assertions.append(_a("fail" if priced else "pass",
                         "Every selected service has a price" if not priced else "No price set for: " + ", ".join(priced), "priced"))
    # dates
    need = max((library["modules"][k]["to_month"] for k in fields["modules"] if k != ADVISORY), default=0)
    assertions.append(_a("pass" if term >= need else "fail",
                         f"Term of {term} months covers the selected services" if term >= need
                         else f"Term of {term} months is shorter than the {need} months the selected services need", "term_covers"))
    end = term_end(start, term)
    last_month_idx = start.year * 12 + (start.month - 1) + term - 1
    ly, lm = last_month_idx // 12, last_month_idx % 12 + 1
    expected = dt.date(ly, lm, calendar.monthrange(ly, lm)[1])
    assertions.append(_a("pass" if end == expected else "fail",
                         f"Term runs {long_date(start)} through {long_date(end)}", "term_dates"))
    assertions.append(_a("pass", "Service dependencies are satisfied", "deps"))

    # narrative
    nar = fields.get("narrative")
    narrative_ok = False
    if not nar:
        assertions.append(_a("warn", "Purpose and objectives are not drafted yet; the document will carry flagged placeholders", "narrative"))
    elif nar.get("key") != narrative_key(fields):
        assertions.append(_a("warn", "Answers changed since the purpose and objectives were drafted; draft them again", "narrative"))
    else:
        bad = check_narrative(list(nar["purpose"]) + list(nar.get("objectives") or []))
        narrative_ok = not bad
        assertions.append(_a("pass" if not bad else "fail",
                             "Drafted purpose and objectives pass the guardrails (no numbers, no guarantees)" if not bad
                             else "Drafted text fails the guardrails: " + "; ".join(bad), "narrative"))
    if not fields["context"] and not fields["objectives"]:
        assertions.append(_a("warn", "No context or objectives captured from the client", "context"))

    sections = build_sections(fields, library, fees, params, narrative_ok)
    summary = summary_rows(fields, library, fees, params)

    # placeholders must all be resolved
    texts = [r[1] for r in summary]
    for s in sections:
        texts.append(s["title"])
        for b in s["blocks"]:
            if "x" in b:
                texts.append(b["x"])
            for row in b.get("rows") or []:
                texts += [str(c) for c in row]
    unresolved = sorted({m for t in texts for m in PLACEHOLDER_RE.findall(t)})
    assertions.append(_a("fail" if unresolved else "pass",
                         "No unresolved placeholders" if not unresolved else "Unresolved placeholders: " + ", ".join(unresolved), "placeholders"))
    assertions.append(_a("pass", f"Fixed clauses come from library version {library.get('version', 1)}", "library_version"))

    ready = not any(a["level"] == "fail" for a in assertions)
    return {"ready": ready, "fields": fields, "missing": [], "assertions": assertions, "fees": fees, "params": params,
            "summary": summary, "sections": sections, "library_version": library.get("version", 1)}
