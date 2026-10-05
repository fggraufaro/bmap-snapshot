"""PostgREST access to the private public.engagement_* tables (service key only).

RLS is enabled on all of them with no policies, and anon/authenticated have no
privileges, so only this service-role path can read clause text, prices or
client briefs.
"""

import datetime as dt

import requests

from ingestion.supabase_client import SUPA_KEY, SUPA_URL

TIMEOUT = 20


class StoreError(Exception):
    pass


def _h(extra=None):
    h = {"apikey": SUPA_KEY, "Authorization": f"Bearer {SUPA_KEY}", "Content-Type": "application/json",
         "Prefer": "return=representation"}
    if extra:
        h.update(extra)
    return h


def _req(method, table, params=None, body=None, headers=None):
    r = requests.request(method, f"{SUPA_URL}/rest/v1/{table}", params=params, json=body,
                         headers=_h(headers), timeout=TIMEOUT)
    if not r.ok:
        raise StoreError(f"{table}: {r.status_code} {r.text[:200]}")
    return r.json() if r.content else []


def _now():
    return dt.datetime.now(dt.timezone.utc).isoformat()


# ── library ────────────────────────────────────────────────────────────
def get_library():
    clauses = _req("GET", "engagement_clauses", {"select": "section,seq,version,style,body,when_modules", "active": "eq.true",
                                                 "order": "section.asc,seq.asc"})
    modules = _req("GET", "engagement_modules", {"select": "*", "active": "eq.true"})
    texts = _req("GET", "engagement_module_text", {"select": "module_key,seq,style,body", "order": "module_key.asc,seq.asc"})
    lib = {"version": max([c["version"] for c in clauses] or [1]), "clauses": {}, "modules": {}, "module_text": {}}
    for c in clauses:
        lib["clauses"].setdefault(c["section"], []).append(
            {"seq": c["seq"], "style": c["style"], "body": c["body"], "when_modules": c["when_modules"] or []})
    for m in modules:
        lib["modules"][m["module_key"]] = {k: m[k] for k in ("name", "fee_label", "from_month", "to_month", "depends_on",
                                                             "activities", "tail_activities", "deliverable", "sort")}
        lib["modules"][m["module_key"]]["price"] = int(float(m["price"]))
    for t in texts:
        lib["module_text"].setdefault(t["module_key"], []).append({"seq": t["seq"], "style": t["style"], "body": t["body"]})
    return lib


# ── engagements ────────────────────────────────────────────────────────
def log_event(engagement_id, event, actor=None, detail=None):
    _req("POST", "engagement_events", body=[{"engagement_id": engagement_id, "event": event, "actor": actor, "detail": detail}])


def create_engagement(client_name, fields, actor=None):
    row = _req("POST", "engagements", body=[{"client_name": client_name}])[0]
    _req("POST", "engagement_versions", body=[{"engagement_id": row["id"], "version": 1, "fields": fields,
                                               "transcript": [], "change_log": []}])
    log_event(row["id"], "created", actor, {"client_name": client_name})
    return row


def list_engagements():
    return _req("GET", "engagements", {"select": "id,client_name,status,current_version,approved_version,approved_by,approved_at,updated_at",
                                       "order": "updated_at.desc", "limit": "100"})


def get_engagement(eid):
    rows = _req("GET", "engagements", {"select": "*", "id": f"eq.{eid}"})
    if not rows:
        return None
    eng = rows[0]
    ver = _req("GET", "engagement_versions", {"select": "*", "engagement_id": f"eq.{eid}",
                                              "version": f"eq.{eng['current_version']}"})[0]
    drafts = _req("GET", "engagement_drafts", {"select": "id,brief_version,library_version,status,assertions,reviewed_by,reviewed_at,note,created_at,sha256",
                                               "engagement_id": f"eq.{eid}", "order": "created_at.desc"})
    return {"engagement": eng, "version": ver, "drafts": drafts}


def save_fields(eid, fields, transcript, change_log, changed):
    """If fields changed, write a new version; otherwise only update the transcript on the current version."""
    eng = _req("GET", "engagements", {"select": "current_version", "id": f"eq.{eid}"})[0]
    cur = eng["current_version"]
    if changed:
        cur += 1
        _req("POST", "engagement_versions", body=[{"engagement_id": eid, "version": cur, "fields": fields,
                                                   "transcript": transcript, "change_log": change_log}])
        _req("PATCH", "engagements", {"id": f"eq.{eid}"}, {"current_version": cur, "updated_at": _now()})
    else:
        _req("PATCH", "engagement_versions", {"engagement_id": f"eq.{eid}", "version": f"eq.{cur}"},
             {"transcript": transcript, "change_log": change_log})
        _req("PATCH", "engagements", {"id": f"eq.{eid}"}, {"updated_at": _now()})
    return cur


def set_current_fields(eid, version, fields, transcript):
    _req("PATCH", "engagement_versions", {"engagement_id": f"eq.{eid}", "version": f"eq.{version}"},
         {"fields": fields, "transcript": transcript})
    _req("PATCH", "engagements", {"id": f"eq.{eid}"}, {"updated_at": _now()})


# ── drafts and sign-off ────────────────────────────────────────────────
def add_draft(eid, brief_version, library_version, assertions, narrative, docx_b64, sha256):
    row = _req("POST", "engagement_drafts", body=[{
        "engagement_id": eid, "brief_version": brief_version, "library_version": library_version,
        "assertions": assertions, "narrative": narrative, "docx_b64": docx_b64, "sha256": sha256}])[0]
    row.pop("docx_b64", None)
    return row


def get_draft(draft_id, with_docx=False):
    cols = "*" if with_docx else "id,engagement_id,brief_version,library_version,status,assertions,reviewed_by,reviewed_at,note,created_at,sha256"
    rows = _req("GET", "engagement_drafts", {"select": cols, "id": f"eq.{draft_id}"})
    return rows[0] if rows else None


def set_draft_status(draft_id, status, actor, note=None):
    body = {"status": status}
    if status == "reviewed":
        body.update({"reviewed_by": actor, "reviewed_at": _now()})
    if note:
        body["note"] = note
    return _req("PATCH", "engagement_drafts", {"id": f"eq.{draft_id}"}, body)[0]


def supersede_other_drafts(eid, keep_id):
    _req("PATCH", "engagement_drafts", {"engagement_id": f"eq.{eid}", "id": f"neq.{keep_id}", "status": "in.(draft,reviewed,sent)"},
         {"status": "superseded"})


def set_engagement_status(eid, status, approved_version=None, approved_by=None):
    body = {"status": status, "updated_at": _now()}
    if status == "approved":
        body.update({"approved_version": approved_version, "approved_by": approved_by, "approved_at": _now()})
    return _req("PATCH", "engagements", {"id": f"eq.{eid}"}, body)[0]
