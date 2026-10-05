"""Flask blueprint for the Engagement Builder (admin-only, Command Center login).

register(app, require_session) mounts everything under /engagement. The caller
passes its own require_session decorator so these routes share the Command
Center's password and token, nothing new is exposed to bank users.
"""

import base64
import datetime as dt
import hashlib
import io

from flask import Blueprint, jsonify, request, send_file

from . import llm, logic
from . import store as _store
from .docxgen import make_docx

# swappable in tests
store = _store
complete = None  # None -> llm.default_complete

EDITABLE = ("client_name", "client_short", "sponsor", "modules", "start", "term", "third_party", "context",
            "objectives", "scope_overrides", "draft_date", "signers")
TRANSITIONS = {"draft": ["reviewed"], "reviewed": ["sent"], "sent": ["approved"]}


def _err(msg, code=400, **extra):
    return jsonify(dict(error=msg, **extra)), code


def _now():
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")


def _preview(fields, library):
    plan = logic.build_plan(fields, library)
    out = {"ready": plan["ready"], "missing": plan.get("missing", []), "assertions": plan["assertions"]}
    if "sections" in plan:
        out.update({"summary": plan["summary"], "fees": plan["fees"], "sections": plan["sections"]})
    return out


def _diff(old, new):
    keys = sorted(k for k in set(old) | set(new) if old.get(k) != new.get(k) and k != "narrative")
    return keys


def _persist(eid, eng, old_fields, new_fields, transcript, by):
    """Save fields; a change creates a new version and reopens an already-reviewed engagement."""
    changed = _diff(old_fields, new_fields)
    log = list(eng["version"].get("change_log") or [])
    if changed:
        log.append({"at": _now(), "by": by, "fields": changed})
    ver = store.save_fields(eid, new_fields, transcript, log, bool(changed))
    if changed and eng["engagement"]["status"] != "draft":
        store.set_engagement_status(eid, "draft")
        store.log_event(eid, "reopened", by, {"fields": changed})
    return ver, changed


def register(app, require_session):
    bp = Blueprint("engagement", __name__, url_prefix="/engagement")

    def route(rule, methods):
        def deco(fn):
            @bp.route(rule, methods=list(methods) + ["OPTIONS"])
            @require_session
            def wrapper(*a, **kw):
                if request.method == "OPTIONS":
                    return jsonify({})
                try:
                    return fn(*a, **kw)
                except store.StoreError as e:
                    return _err("The database request failed: " + str(e)[:200], 502)
                except llm.LLMUnavailable as e:
                    return _err(str(e), 503)
            wrapper.__name__ = fn.__name__
            return wrapper
        return deco

    def _load(eid):
        eng = store.get_engagement(eid)
        if not eng:
            return None, None
        return eng, store.get_library()

    @route("/catalog", ["GET"])
    def catalog():
        lib = store.get_library()
        mods = [{"key": k, "name": lib["modules"][k]["name"], "depends_on": lib["modules"][k]["depends_on"],
                 "min_term": lib["modules"][k]["to_month"]} for k in logic.choosable_modules(lib)]
        return jsonify({"library_version": lib["version"], "modules": mods, "min_term": logic.MIN_TERM, "max_term": logic.MAX_TERM})

    @route("/list", ["GET"])
    def list_():
        return jsonify(store.list_engagements())

    @route("/create", ["POST"])
    def create():
        b = request.get_json(force=True, silent=True) or {}
        name = str(b.get("client_name") or "").strip()
        if not name:
            return _err("client_name is required")
        lib = store.get_library()
        fields, _ = logic.clean_fields({"client_name": name, "client_short": b.get("client_short"),
                                        "sponsor": b.get("sponsor"), "modules": []}, lib)
        row = store.create_engagement(name, fields, actor=str(b.get("actor") or "operator")[:60])
        return jsonify({"id": row["id"]})

    @route("/<eid>", ["GET"])
    def get_(eid):
        eng, lib = _load(eid)
        if not eng:
            return _err("not found", 404)
        return jsonify({"engagement": eng["engagement"], "fields": eng["version"]["fields"],
                        "transcript": eng["version"]["transcript"], "change_log": eng["version"]["change_log"],
                        "drafts": eng["drafts"], "preview": _preview(eng["version"]["fields"], lib)})

    @route("/<eid>/chat", ["POST"])
    def chat(eid):
        msg = str((request.get_json(force=True, silent=True) or {}).get("message") or "").strip()[:2000]
        if not msg:
            return _err("message is required")
        eng, lib = _load(eid)
        if not eng:
            return _err("not found", 404)
        fields = eng["version"]["fields"]
        transcript = list(eng["version"]["transcript"] or [])
        history = [h for h in transcript if h.get("w") in ("ai", "me")]
        res = llm.chat_turn(fields, history, msg, lib, complete=complete)
        transcript.append({"w": "me", "t": msg, "at": _now()})
        transcript.append({"w": "ai", "t": res["reply"], "at": _now()})
        for n in res["notes"] + res["flags"]:
            transcript.append({"w": "sys", "t": n, "at": _now()})
        ver, changed = _persist(eid, eng, fields, res["fields"], transcript, "chat")
        return jsonify({"ok": res["ok"], "reply": res["reply"], "notes": res["notes"], "flags": res["flags"],
                        "suggestions": res["suggestions"], "quick_replies": res["quick_replies"],
                        "version": ver, "changed": changed, "fields": res["fields"],
                        "preview": _preview(res["fields"], lib)})

    @route("/<eid>/suggestion", ["POST"])
    def accept(eid):
        b = request.get_json(force=True, silent=True) or {}
        sug = {"field": b.get("field"), "value": b.get("value")}
        eng, lib = _load(eid)
        if not eng:
            return _err("not found", 404)
        if not llm.valid_suggestion(sug, lib):
            return _err("That suggestion is not valid")
        fields = eng["version"]["fields"]
        new_fields, notes = llm.apply_updates(fields, {sug["field"]: sug["value"]}, lib)
        transcript = list(eng["version"]["transcript"] or [])
        transcript.append({"w": "sys", "t": "Accepted: " + llm.suggestion_label(sug, lib) + ".", "at": _now()})
        for n in notes:
            transcript.append({"w": "sys", "t": n, "at": _now()})
        ver, changed = _persist(eid, eng, fields, new_fields, transcript, "accept")
        return jsonify({"version": ver, "fields": new_fields, "notes": notes, "preview": _preview(new_fields, lib)})

    @route("/<eid>/fields", ["POST"])
    def edit_fields(eid):
        updates = (request.get_json(force=True, silent=True) or {}).get("updates") or {}
        eng, lib = _load(eid)
        if not eng:
            return _err("not found", 404)
        fields = dict(eng["version"]["fields"])
        for k in EDITABLE:
            if k in updates:
                fields[k] = updates[k]
        new_fields, errors = logic.clean_fields(fields, lib)
        new_fields["narrative"] = eng["version"]["fields"].get("narrative")
        transcript = list(eng["version"]["transcript"] or [])
        transcript.append({"w": "sys", "t": "Operator edited: " + ", ".join(k for k in EDITABLE if k in updates), "at": _now()})
        ver, changed = _persist(eid, eng, eng["version"]["fields"], new_fields, transcript, "operator")
        return jsonify({"version": ver, "fields": new_fields, "errors": errors, "preview": _preview(new_fields, lib)})

    @route("/<eid>/narrative", ["POST"])
    def narrative(eid):
        eng, lib = _load(eid)
        if not eng:
            return _err("not found", 404)
        fields = eng["version"]["fields"]
        try:
            nar = llm.draft_narrative(fields, complete=complete)
        except llm.NarrativeError as e:
            return _err(str(e), 422)
        new_fields = dict(fields, narrative=nar)
        transcript = list(eng["version"]["transcript"] or [])
        transcript.append({"w": "sys", "t": "Purpose and objectives drafted by Claude; needs review.", "at": _now()})
        # narrative is derived, not an operator-facing field change: update the current version in place
        store.set_current_fields(eid, eng["engagement"]["current_version"], new_fields, transcript)
        return jsonify({"narrative": nar, "preview": _preview(new_fields, lib)})

    @route("/<eid>/generate", ["POST"])
    def generate(eid):
        eng, lib = _load(eid)
        if not eng:
            return _err("not found", 404)
        fields = eng["version"]["fields"]
        plan = logic.build_plan(fields, lib)
        if not plan["ready"]:
            return _err("The draft cannot be generated yet", 422, missing=plan.get("missing", []),
                        assertions=plan["assertions"])
        version = eng["engagement"]["current_version"]
        data = make_docx(plan, version)
        sha = hashlib.sha256(data).hexdigest()
        row = store.add_draft(eid, version, plan["library_version"], plan["assertions"], fields.get("narrative"),
                              base64.b64encode(data).decode(), sha)
        store.log_event(eid, "draft_generated", "operator", {"draft": row["id"], "brief_version": version, "sha256": sha})
        return jsonify({"draft": row, "assertions": plan["assertions"]})

    @route("/<eid>/draft/<did>/docx", ["GET"])
    def download(eid, did):
        d = store.get_draft(did, with_docx=True)
        if not d or str(d["engagement_id"]) != eid or not d.get("docx_b64"):
            return _err("not found", 404)
        data = base64.b64decode(d["docx_b64"])
        eng = store.get_engagement(eid)
        name = (eng["engagement"]["client_name"] if eng else "engagement").replace(" ", "_")
        return send_file(io.BytesIO(data), as_attachment=True,
                         download_name=f"SOW_{name}_v{d['brief_version']}_{d['status']}.docx",
                         mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    @route("/<eid>/draft/<did>/status", ["POST"])
    def draft_status(eid, did):
        b = request.get_json(force=True, silent=True) or {}
        to, actor = b.get("status"), str(b.get("actor") or "").strip()[:80]
        note = str(b.get("note") or "").strip()[:300] or None
        d = store.get_draft(did)
        eng = store.get_engagement(eid)
        if not d or not eng or str(d["engagement_id"]) != eid:
            return _err("not found", 404)
        if to not in TRANSITIONS.get(d["status"], []):
            return _err(f"A {d['status']} draft cannot move to {to}", 409)
        if not actor:
            return _err("Enter your name; the sign-off log records who did this", 400)
        if d["brief_version"] != eng["engagement"]["current_version"]:
            return _err("The brief changed after this draft was generated. Generate a new draft.", 409)
        if any(a["level"] == "fail" for a in d["assertions"]):
            return _err("This draft has failed checks", 409)
        row = store.set_draft_status(did, to, actor, note)
        if to == "approved":
            store.set_engagement_status(eid, "approved", d["brief_version"], actor)
            store.supersede_other_drafts(eid, did)
        elif to in ("reviewed", "sent"):
            store.set_engagement_status(eid, to)
        store.log_event(eid, "draft_" + to, actor, {"draft": did, "brief_version": d["brief_version"], "note": note})
        return jsonify({"draft": row})

    app.register_blueprint(bp)
    return bp
