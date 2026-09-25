"""
BMAP Pipeline Command Center API -- Railway deployment (separate service
from main.py)
============================================================================
Flask app for triggering and monitoring the data pipeline (ingestion +
rebuild steps) from a simple click-to-run admin page, instead of running
each step by hand through the Supabase SQL editor.

Deliberately a separate service and a separate password from the Hub
(main.py / secure_proxy.py): this panel can trigger real production data
rebuilds (rebuilding branches_master_v2, the competitor systems, and
branch_opportunity_base), so it isn't meant to share the whole team's Hub
login.

Every step runs in a background thread and reports status through
public.pipeline_jobs (Supabase-backed, so job history survives a Railway
restart/redeploy) -- the HTTP request returns almost immediately with a
job_id, and the client polls /status/<job_id>. This matters because a
single step can take many minutes (the tiered competitor rebuild alone runs
10+ minutes), far past what an HTTP request should be left open for.

Endpoints:
  POST /auth/login         { password }  -> { token }
  GET  /steps              Bearer token  -> [{ id, label, kind, last_run }]
  POST /run/<step_id>      Bearer token  -> { job_id }   (runs one step, async)
  POST /run-all            Bearer token  -> { job_id }   (runs the full pipeline in order, async)
  GET  /status/<job_id>    Bearer token  -> pipeline_jobs row
  GET  /jobs               Bearer token  -> last 50 pipeline_jobs rows
  GET  /health             -> { status: ok }

Required Railway env vars:
  SUPABASE_SERVICE_KEY     -- same one the ingestion scripts already use
  PIPELINE_ADMIN_PASSWORD  -- separate from the Hub's HUB_ACCESS_PASSWORD
  PIPELINE_SESSION_SECRET  -- any long random string, HMAC-signs session tokens
  ALLOWED_ORIGIN           -- the command-center page's origin (GitHub Pages)

Generate PIPELINE_SESSION_SECRET with:
  python -c "import secrets; print(secrets.token_hex(32))"
"""

import base64
import hashlib
import hmac
import os
import threading
import time
import traceback
from datetime import datetime, timezone
from functools import wraps

import requests as http
from flask import Flask, jsonify, request
from flask_cors import CORS

from ingestion.pipeline_steps import ALL_STEPS, RUN_ALL_ORDER, STEP_BY_ID
from ingestion.supabase_client import SUPA_KEY, SUPA_URL

app = Flask(__name__)

ALLOWED_ORIGIN = os.environ.get("ALLOWED_ORIGIN", "https://fggraufaro.github.io")
CORS(app, origins=[ALLOWED_ORIGIN])

PIPELINE_PASSWORD = os.environ.get("PIPELINE_ADMIN_PASSWORD", "")
SESSION_SECRET = os.environ.get("PIPELINE_SESSION_SECRET", "")
SESSION_TTL_SECONDS = 12 * 60 * 60  # 12 hours

# Same fail-fast rationale as secure_proxy.py: an empty secret/password
# would either forge-able tokens or silently reject every login.
if not SESSION_SECRET:
    raise RuntimeError(
        "PIPELINE_SESSION_SECRET env var is not set. Refusing to start: an "
        "empty HMAC secret would let anyone forge valid session tokens. "
        "Generate one with: python -c \"import secrets; print(secrets.token_hex(32))\""
    )
if not PIPELINE_PASSWORD:
    raise RuntimeError(
        "PIPELINE_ADMIN_PASSWORD env var is not set. Refusing to start: with "
        "no password configured, /auth/login would reject every attempt anyway."
    )

_login_attempts = {}  # ip -> [timestamps]; in-memory, resets on redeploy -- fine, single small team
LOGIN_MAX_ATTEMPTS = 8
LOGIN_WINDOW_SECONDS = 5 * 60


# ── Session token: HMAC-signed, same scheme as secure_proxy.py ─────────
def _make_token(subject="pipeline"):
    exp = int(time.time()) + SESSION_TTL_SECONDS
    payload = f"{subject}:{exp}"
    sig = hmac.new(SESSION_SECRET.encode(), payload.encode(), hashlib.sha256).hexdigest()
    raw = f"{payload}:{sig}"
    return base64.urlsafe_b64encode(raw.encode()).decode()


def _verify_token(token):
    try:
        raw = base64.urlsafe_b64decode(token.encode()).decode()
        subject, exp, sig = raw.split(":")
        payload = f"{subject}:{exp}"
        expected = hmac.new(SESSION_SECRET.encode(), payload.encode(), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(sig, expected):
            return False
        if int(exp) < time.time():
            return False
        return True
    except Exception:
        return False


def require_session(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if request.method == "OPTIONS":
            return fn(*args, **kwargs)
        auth = request.headers.get("Authorization", "")
        token = auth.replace("Bearer ", "").strip()
        if not token or not _verify_token(token):
            return jsonify({"error": "unauthorized"}), 401
        return fn(*args, **kwargs)
    return wrapper


def _cors_headers(resp):
    resp.headers["Access-Control-Allow-Origin"] = ALLOWED_ORIGIN
    resp.headers["Access-Control-Allow-Headers"] = "Authorization, Content-Type"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    return resp


@app.after_request
def _apply_cors(resp):
    return _cors_headers(resp)


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "BMAP Pipeline Command Center"})


@app.route("/auth/login", methods=["POST", "OPTIONS"])
def login():
    if request.method == "OPTIONS":
        return _cors_headers(jsonify({}))

    ip = request.headers.get("X-Forwarded-For", request.remote_addr) or "unknown"
    now = time.time()
    attempts = [t for t in _login_attempts.get(ip, []) if now - t < LOGIN_WINDOW_SECONDS]
    if len(attempts) >= LOGIN_MAX_ATTEMPTS:
        return jsonify({"error": "too many attempts -- try again later"}), 429

    body = request.get_json(force=True, silent=True) or {}
    password = (body.get("password") or "").strip()

    attempts.append(now)
    _login_attempts[ip] = attempts

    if not hmac.compare_digest(password, PIPELINE_PASSWORD):
        return jsonify({"error": "incorrect password"}), 401

    _login_attempts[ip] = []
    token = _make_token()
    return jsonify({"token": token, "expires_in": SESSION_TTL_SECONDS})


# ── Job tracking -- Supabase-backed so history survives a redeploy ─────
def _jobs_headers():
    return {"apikey": SUPA_KEY, "Authorization": f"Bearer {SUPA_KEY}", "Content-Type": "application/json"}


def _job_create(step_id):
    url = f"{SUPA_URL}/rest/v1/pipeline_jobs"
    headers = _jobs_headers()
    headers["Prefer"] = "return=representation"
    r = http.post(url, headers=headers, json={"step_id": step_id, "status": "pending"}, timeout=15)
    r.raise_for_status()
    return r.json()[0]["id"]


def _job_write(job_id, **fields):
    url = f"{SUPA_URL}/rest/v1/pipeline_jobs?id=eq.{job_id}"
    headers = _jobs_headers()
    headers["Prefer"] = "return=minimal"
    try:
        http.patch(url, headers=headers, json=fields, timeout=15)
    except Exception as e:
        print(f"[pipeline] failed to write job status for {job_id}: {e}")


def _job_read(job_id):
    url = f"{SUPA_URL}/rest/v1/pipeline_jobs?id=eq.{job_id}&select=*"
    r = http.get(url, headers=_jobs_headers(), timeout=15)
    r.raise_for_status()
    rows = r.json()
    return rows[0] if rows else None


def _run_step_job(job_id, step_id):
    try:
        _job_write(job_id, status="running")
        STEP_BY_ID[step_id]["fn"]()
        _job_write(job_id, status="done", finished_at=datetime.now(timezone.utc).isoformat())
        print(f"[pipeline] {step_id} ({job_id}) done")
    except Exception as e:
        tb = traceback.format_exc()
        print(f"[pipeline] {step_id} ({job_id}) FAILED: {e}\n{tb}")
        _job_write(job_id, status="error", error_message=str(e)[:2000],
                   finished_at=datetime.now(timezone.utc).isoformat())


def _run_all_job(job_id):
    log_lines = []
    try:
        _job_write(job_id, status="running")
        for step_id in RUN_ALL_ORDER:
            log_lines.append(f"-> {step_id}")
            _job_write(job_id, log="\n".join(log_lines))
            STEP_BY_ID[step_id]["fn"]()
            log_lines[-1] += " ok"
            _job_write(job_id, log="\n".join(log_lines))
        _job_write(job_id, status="done", finished_at=datetime.now(timezone.utc).isoformat())
        print(f"[pipeline] run-all ({job_id}) done")
    except Exception as e:
        tb = traceback.format_exc()
        print(f"[pipeline] run-all ({job_id}) FAILED: {e}\n{tb}")
        if log_lines:
            log_lines[-1] += f" FAILED: {e}"
        _job_write(job_id, status="error", error_message=str(e)[:2000],
                   log="\n".join(log_lines), finished_at=datetime.now(timezone.utc).isoformat())


@app.route("/steps", methods=["GET", "OPTIONS"])
@require_session
def list_steps():
    if request.method == "OPTIONS":
        return _cors_headers(jsonify({}))

    out = []
    for s in ALL_STEPS:
        url = (f"{SUPA_URL}/rest/v1/pipeline_jobs?step_id=eq.{s['id']}"
               f"&select=status,started_at,finished_at,error_message&order=started_at.desc&limit=1")
        try:
            r = http.get(url, headers=_jobs_headers(), timeout=15)
            last = r.json()[0] if r.ok and r.json() else None
        except Exception:
            last = None
        out.append({"id": s["id"], "label": s["label"], "kind": s["kind"], "last_run": last})
    return jsonify(out)


@app.route("/run/<step_id>", methods=["POST", "OPTIONS"])
@require_session
def run_step(step_id):
    if request.method == "OPTIONS":
        return _cors_headers(jsonify({}))
    if step_id not in STEP_BY_ID:
        return jsonify({"error": f"unknown step '{step_id}'"}), 404

    try:
        job_id = _job_create(step_id)
    except Exception as e:
        return jsonify({"error": f"could not create job: {e}"}), 500

    thread = threading.Thread(target=_run_step_job, args=(job_id, step_id), daemon=True)
    thread.start()
    print(f"[pipeline] started {step_id} ({job_id})")
    return jsonify({"job_id": job_id}), 202


@app.route("/run-all", methods=["POST", "OPTIONS"])
@require_session
def run_all():
    if request.method == "OPTIONS":
        return _cors_headers(jsonify({}))

    try:
        job_id = _job_create("run_all")
    except Exception as e:
        return jsonify({"error": f"could not create job: {e}"}), 500

    thread = threading.Thread(target=_run_all_job, args=(job_id,), daemon=True)
    thread.start()
    print(f"[pipeline] started run-all ({job_id})")
    return jsonify({"job_id": job_id}), 202


@app.route("/status/<job_id>", methods=["GET", "OPTIONS"])
@require_session
def status(job_id):
    if request.method == "OPTIONS":
        return _cors_headers(jsonify({}))
    row = _job_read(job_id)
    if not row:
        return jsonify({"error": "job not found"}), 404
    return jsonify(row)


@app.route("/jobs", methods=["GET", "OPTIONS"])
@require_session
def jobs():
    if request.method == "OPTIONS":
        return _cors_headers(jsonify({}))
    url = f"{SUPA_URL}/rest/v1/pipeline_jobs?select=*&order=started_at.desc&limit=50"
    r = http.get(url, headers=_jobs_headers(), timeout=15)
    return jsonify(r.json()), r.status_code


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8081))
    print(f"BMAP Pipeline Command Center API starting on port {port}")
    app.run(host="0.0.0.0", port=port)
