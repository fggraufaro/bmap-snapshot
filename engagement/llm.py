"""Claude-backed conversation and narrative drafting, with guardrails.

Claude never sees prices and never sets fees, dates or clauses. It proposes
field updates (applied only after validation here) and drafts the tailored
purpose and objectives sections, which must pass logic.check_narrative.

`complete` is injectable (prompt, max_tokens) -> str so tests run without a key.
"""

import json
import os
import re

from . import logic

MODEL = os.environ.get("ENGAGEMENT_MODEL", "claude-sonnet-4-6")
ANTH_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
HISTORY_TURNS = 14


class NarrativeError(Exception):
    pass


class LLMUnavailable(Exception):
    pass


def default_complete(prompt, max_tokens=1500):
    if not ANTH_KEY:
        raise LLMUnavailable("ANTHROPIC_API_KEY is not set on this service")
    try:
        import anthropic
    except ImportError as e:  # pragma: no cover
        raise LLMUnavailable("anthropic package is not installed") from e
    client = anthropic.Anthropic(api_key=ANTH_KEY)
    msg = client.messages.create(model=MODEL, max_tokens=max_tokens,
                                 messages=[{"role": "user", "content": prompt}])
    return "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")


def parse_json(raw):
    raw = (raw or "").strip()
    raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw)
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", raw, flags=re.S)
        if m:
            return json.loads(m.group(0))
        raise


# ── applying proposed updates ──────────────────────────────────────────
def apply_updates(fields, updates, library):
    """Validate and merge an updates dict into fields. Returns (new_fields, notes)."""
    f = json.loads(json.dumps(fields))
    notes = []
    u = updates if isinstance(updates, dict) else {}
    for key, maxlen in (("client_name", 120), ("client_short", 60), ("sponsor", 80)):
        v = u.get(key)
        if isinstance(v, str) and v.strip():
            f[key] = v.strip()[:maxlen]
    if isinstance(u.get("modules"), list):
        wanted = [m for m in u["modules"] if isinstance(m, str) and m in library["modules"] and m != logic.ADVISORY]
        kept, removed = logic.drop_unmet_dependencies(wanted, library)
        f["modules"] = logic.order_modules(kept, library)
        if removed:
            notes.append("Removed because they depend on a step that is not selected: " + ", ".join(removed) + ".")
    s = u.get("start")
    if isinstance(s, str) and logic.START_RE.match(s):
        f["start"] = s
    t = u.get("term")
    if isinstance(t, int) and not isinstance(t, bool) and logic.MIN_TERM <= t <= logic.MAX_TERM:
        f["term"] = t
    if u.get("third_party") in ("pass", "direct"):
        f["third_party"] = u["third_party"]
    for key in ("context", "objectives"):
        if isinstance(u.get(key), list):
            merged = list(f.get(key) or [])
            for x in u[key]:
                if isinstance(x, str) and x.strip() and len(x.strip()) <= 160 \
                        and x.strip().lower() not in [y.lower() for y in merged] and len(merged) < 8:
                    merged.append(x.strip())
            f[key] = merged
    return f, notes


def valid_suggestion(s, library):
    if not isinstance(s, dict):
        return False
    field, v = s.get("field"), s.get("value")
    if field == "modules":
        return isinstance(v, list) and any(isinstance(m, str) and m in library["modules"] and m != logic.ADVISORY for m in v)
    if field == "start":
        return isinstance(v, str) and bool(logic.START_RE.match(v))
    if field == "term":
        return isinstance(v, int) and not isinstance(v, bool) and logic.MIN_TERM <= v <= logic.MAX_TERM
    if field == "third_party":
        return v in ("pass", "direct")
    return False


def suggestion_label(s, library):
    if s["field"] == "modules":
        return "; ".join(library["modules"][m]["name"] for m in s["value"] if m in library["modules"])
    if s["field"] == "start":
        y, m, _ = s["value"].split("-")
        import calendar
        return f"Start {calendar.month_name[int(m)]} {y}"
    if s["field"] == "term":
        return f"{s['value']} months"
    return ("Third-party costs passed through by Verlocity" if s["value"] == "pass"
            else "The bank contracts providers directly")


# ── conversation ───────────────────────────────────────────────────────
def _instructions(library):
    catalog = []
    for k in logic.choosable_modules(library):
        m = library["modules"][k]
        dep = f", requires {m['depends_on']}" if m.get("depends_on") else ""
        catalog.append(f"{k} = {m['name']}{dep}; needs at least {m['to_month']} months")
    return (
        "You are the intake assistant for Verlocity, a deposit-growth advisory firm for community banks. "
        "You are in a live conversation where an operator is scoping a statement of work with a bank CMO. "
        "Understand the bank's situation and goals and fill in the brief. Be warm and concise: one to three sentences per reply.\n"
        "ADAPT your questions to what has been said. Pick the single most useful next question. Examples: if they mention a new "
        "digital platform, ask what is blocking funded accounts; if they mention new markets, ask which and how soon; if they "
        "are short on time or budget, favor fewer services. Never ask something already answered. Ask exactly ONE question per "
        "reply, except when the brief is complete and you are wrapping up.\n"
        "SERVICES (module keys): " + " | ".join(catalog) + ". Strategic advisory is always included; never list it.\n"
        f"BRIEF FIELDS: client_name; client_short; sponsor (the bank's executive sponsor, if named); modules (list of keys above); "
        f"start (first day of a month, YYYY-MM-01); term (whole months, {logic.MIN_TERM} to {logic.MAX_TERM}); context (facts about "
        "what is already in play at the bank, short phrases); objectives (what the CMO wants, short phrases); third_party "
        "('pass' = Verlocity bills approved third-party costs as pass-through; 'direct' = the bank contracts providers directly).\n"
        "RULES: Put a value in updates ONLY if the CMO clearly said it. If you are proposing something they did not say, put it in "
        "suggestions with a one-line reason; the operator accepts or declines it. Never invent facts, numbers, bank details or "
        "prices, and never quote or discuss fees. If a request conflicts with the field rules (for example a step that needs more "
        "months than the term), say so and offer an option. For context and objectives return only NEW items from this message. "
        "For modules return the full intended list.\n"
        'OUTPUT: JSON only, exactly: {"reply": string, "updates": {"client_name": string|null, "client_short": string|null, '
        '"sponsor": string|null, "modules": string[]|null, "start": string|null, "term": number|null, "context": string[]|null, '
        '"objectives": string[]|null, "third_party": "pass"|"direct"|null}, "suggestions": [{"field": "modules"|"start"|"term"|'
        '"third_party", "value": any, "why": string}], "quick_replies": string[] (up to 3 short likely answers to your '
        'question), "flags": string[]}'
    )


def chat_turn(fields, history, message, library, complete=None, today=None):
    """One conversational turn. history = [{"w": "ai"|"me", "t": text}, ...] (prior turns,
    not including `message`). Returns a dict; never raises on a bad model reply."""
    complete = complete or default_complete
    import datetime as dt
    brief = {k: fields.get(k) for k in ("client_name", "client_short", "sponsor", "modules", "start", "term",
                                        "context", "objectives", "third_party")}
    hist = "\n".join(("ASSISTANT: " if h["w"] == "ai" else "CMO: ") + h["t"] for h in history[-HISTORY_TURNS:])
    prompt = (_instructions(library) + f"\n\nToday is {today or dt.date.today().isoformat()}.\n\nCURRENT BRIEF (JSON):\n"
              + json.dumps(brief) + "\n\nCONVERSATION SO FAR:\n" + (hist or "(none)")
              + "\n\nThe CMO now says: " + message + "\n\nRespond with the JSON object only.")
    try:
        data = parse_json(complete(prompt, 1200))
        if not isinstance(data, dict) or not isinstance(data.get("reply"), str):
            raise ValueError("shape")
    except LLMUnavailable:
        raise
    except Exception:
        return {"ok": False, "reply": "I could not read my own reply just now. Please send that again.",
                "fields": fields, "notes": [], "suggestions": [], "quick_replies": [], "flags": []}
    new_fields, notes = apply_updates(fields, data.get("updates"), library)
    sugs = [dict(s, label=suggestion_label(s, library)) for s in (data.get("suggestions") or [])
            if valid_suggestion(s, library)][:3]
    for s in sugs:
        s["why"] = str(s.get("why") or "")[:200]
    qr = [q for q in (data.get("quick_replies") or []) if isinstance(q, str) and len(q) < 90][:3]
    flags = [str(x)[:300] for x in (data.get("flags") or []) if isinstance(x, str)][:2]
    return {"ok": True, "reply": data["reply"][:700], "fields": new_fields, "notes": notes,
            "suggestions": sugs, "quick_replies": qr, "flags": flags}


# ── narrative ──────────────────────────────────────────────────────────
def draft_narrative(fields, complete=None):
    """Draft the tailored purpose and objectives. Raises NarrativeError if the
    guardrails fail twice. Returns {"purpose": [paragraphs], "objectives": [...], "key": ...}."""
    complete = complete or default_complete
    ctx, obj = fields.get("context") or [], fields.get("objectives") or []
    if not ctx and not obj:
        raise NarrativeError("Nothing to draft from: no context or objectives captured yet.")
    base = (
        "Write sections of a statement of work for a bank engagement. Use ONLY these inputs.\n"
        f"Client: {fields.get('client_short') or fields.get('client_name')}\n"
        f"What is in play at the bank: {json.dumps(ctx)}\n"
        f"What the CMO wants: {json.dumps(obj)}\n\n"
        'Return JSON only: {"purpose": [string, ...], "objectives": [string, ...]}. "purpose" is two or three short paragraphs '
        "explaining why the client is engaging Verlocity, in the client's own context, ending with a sentence that Verlocity will "
        "serve as a collaborative strategic partner across these connected priorities. \"objectives\" has one sentence per input "
        "objective, in the same order (an empty list if there are none). RULES: use only the inputs; write no numbers or digits "
        "anywhere (say \"additional markets\", not a count); invent no names, facts, results or commitments; no guarantees or "
        "promises; neutral professional tone; refer to the client by its short name and to \"Verlocity\"."
    )
    problems = ""
    for attempt in range(2):
        raw = complete(base + problems, 1200)
        try:
            d = parse_json(raw)
            purpose = [p.strip() for p in d.get("purpose", []) if isinstance(p, str) and p.strip()][:4]
            objs = [o.strip() for o in d.get("objectives", []) if isinstance(o, str) and o.strip()][:8]
            if not purpose:
                raise ValueError("no purpose")
        except Exception:
            problems = "\n\nYour previous reply was not valid JSON in the required shape. Return the JSON object only."
            continue
        fails = logic.check_narrative(purpose + objs)
        if not fails:
            return {"purpose": purpose, "objectives": objs, "key": logic.narrative_key(fields)}
        problems = "\n\nYour previous draft failed these checks: " + "; ".join(fails) + ". Rewrite it so none of them apply."
    raise NarrativeError("The drafted text did not pass the guardrails after two attempts. " + problems.strip())
