"""
secure_proxy.py — Verlocity Hub auth + data proxy
====================================================
Replaces the pattern where the Supabase anon key and Anthropic key sat
in plain text inside context-generator.html. Now:

  - The browser never sees a Supabase key or an Anthropic key.
  - The browser gets a short-lived signed session token after a
    password check, and sends that token on every request.
  - This module validates the token, then does the actual Supabase /
    Anthropic call server-side using the SERVICE ROLE key (never
    shipped to the client) and returns just the JSON payload.
  - Only an explicit allowlist of tables/views can be queried — no
    arbitrary table access even with a valid session token.

Wire into main.py with:

    from secure_proxy import secure_proxy_bp
    app.register_blueprint(secure_proxy_bp)

Required Railway env vars (Settings → Variables):
    SUPABASE_SERVICE_KEY   — Settings → API → service_role (NOT anon)
    ANTHROPIC_API_KEY      — already set for bmap_snapshot.py / bmap_board_brief.py
    HUB_ACCESS_PASSWORD    — the passphrase the team uses to log into the Hub
    SESSION_SECRET         — any long random string, used to sign session tokens
    ALLOWED_ORIGIN         — https://fggraufaro.github.io (locks CORS down from '*')

Generate a SESSION_SECRET quickly with:
    python -c "import secrets; print(secrets.token_hex(32))"
"""

import os
import time
import hmac
import hashlib
import base64
import json
from functools import wraps

import requests
from flask import Blueprint, request, jsonify

secure_proxy_bp = Blueprint("secure_proxy", __name__)

# ── Config ──────────────────────────────────────────────────────
SUPA_URL      = "https://tuiiywphoynbmkxpoyps.supabase.co"
SUPA_SERVICE  = os.environ.get("SUPABASE_SERVICE_KEY", "")
ANTH_KEY      = os.environ.get("ANTHROPIC_API_KEY", "")
HUB_PASSWORD  = os.environ.get("HUB_ACCESS_PASSWORD", "")
SESSION_SECRET = os.environ.get("SESSION_SECRET", "")
ALLOWED_ORIGIN = os.environ.get("ALLOWED_ORIGIN", "https://fggraufaro.github.io")

SESSION_TTL_SECONDS = 12 * 60 * 60  # 12 hours — re-login next day

# Only these tables/views are reachable through the proxy. Anything
# else is refused, even with a valid session token. This mirrors
# exactly what context-generator.html's SCHEMA_MAP + api() calls use,
# verified directly against the live database (not guessed from code
# fragments — two of these live outside 'public' and got this wrong
# on the first pass).
ALLOWED_TABLES = {
    "dim_institutions":                 "ref",
    "bank_website":                     "ref",
    "branch_opportunity_base":          "analytics",
    "branch_target_competitors":        "analytics",
    "bank_financial_snapshot_latest":   "analytics",
    "vw_branch_opportunity_cbsa":       "public",
    "vw_network_top_targets":           "public",
    "vw_prospecting_score":             "public",
    "vw_zip_persona":                   "public",
    "uszips":                           "geo",
}

# Postgres functions the Hub calls via rpc(). All four live in 'public'.
ALLOWED_RPCS = {
    "branches_within_radius",
    "radius_market_summary",
    "radius_opportunity_extremes",
    "radius_zip_detail",
}

# Very small in-memory rate limiter for the login endpoint.
# Resets on redeploy — fine for a small internal team tool.
_login_attempts = {}  # ip -> [timestamps]
LOGIN_MAX_ATTEMPTS = 8
LOGIN_WINDOW_SECONDS = 5 * 60


# ── Session token: HMAC-signed, not a JWT library dependency ──────
def _make_token(subject: str = "hub") -> str:
    exp = int(time.time()) + SESSION_TTL_SECONDS
    payload = f"{subject}:{exp}"
    sig = hmac.new(SESSION_SECRET.encode(), payload.encode(), hashlib.sha256).hexdigest()
    raw = f"{payload}:{sig}"
    return base64.urlsafe_b64encode(raw.encode()).decode()


def _verify_token(token: str) -> bool:
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
        # Preflight requests never carry the Authorization header — let
        # them through untouched so CORS can succeed, then the browser's
        # real request (which does carry the token) hits the auth check.
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


@secure_proxy_bp.after_request
def _apply_cors(resp):
    return _cors_headers(resp)


@secure_proxy_bp.route("/auth/login", methods=["POST", "OPTIONS"])
def login():
    if request.method == "OPTIONS":
        return _cors_headers(jsonify({}))

    ip = request.headers.get("X-Forwarded-For", request.remote_addr) or "unknown"
    now = time.time()
    attempts = [t for t in _login_attempts.get(ip, []) if now - t < LOGIN_WINDOW_SECONDS]
    if len(attempts) >= LOGIN_MAX_ATTEMPTS:
        return jsonify({"error": "too many attempts — try again later"}), 429

    body = request.get_json(force=True, silent=True) or {}
    password = (body.get("password") or "").strip()

    attempts.append(now)
    _login_attempts[ip] = attempts

    if not HUB_PASSWORD or not hmac.compare_digest(password, HUB_PASSWORD):
        return jsonify({"error": "incorrect password"}), 401

    _login_attempts[ip] = []  # reset on success
    token = _make_token()
    return jsonify({"token": token, "expires_in": SESSION_TTL_SECONDS})


@secure_proxy_bp.route("/api/<table>", methods=["GET", "OPTIONS"])
@require_session
def proxy_table(table):
    if request.method == "OPTIONS":
        return _cors_headers(jsonify({}))

    schema = ALLOWED_TABLES.get(table)
    if schema is None:
        return jsonify({"error": f"table '{table}' is not exposed via the proxy"}), 403

    # Forward the querystring as-is (select=, filters, order, limit —
    # these are the same params the Hub already builds client-side).
    qs = request.query_string.decode()
    url = f"{SUPA_URL}/rest/v1/{table}?{qs}"

    try:
        r = requests.get(
            url,
            headers={
                "apikey": SUPA_SERVICE,
                "Authorization": f"Bearer {SUPA_SERVICE}",
                "Accept-Profile": schema,
            },
            timeout=20,
        )
        return jsonify(r.json()), r.status_code
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@secure_proxy_bp.route("/api/rpc/<fn_name>", methods=["POST", "OPTIONS"])
@require_session
def proxy_rpc(fn_name):
    if request.method == "OPTIONS":
        return _cors_headers(jsonify({}))

    if fn_name not in ALLOWED_RPCS:
        return jsonify({"error": f"function '{fn_name}' is not exposed via the proxy"}), 403

    body = request.get_json(force=True, silent=True) or {}
    url = f"{SUPA_URL}/rest/v1/rpc/{fn_name}"

    try:
        r = requests.post(
            url,
            headers={
                "apikey": SUPA_SERVICE,
                "Authorization": f"Bearer {SUPA_SERVICE}",
                "Content-Type": "application/json",
            },
            json=body,
            timeout=20,
        )
        return jsonify(r.json()), r.status_code
    except Exception as e:
        return jsonify({"error": str(e)}), 500
