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
from datetime import datetime, timezone
from functools import wraps
from urllib.parse import quote

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

# Fail fast if these are missing rather than silently running with an empty
# string. An empty SESSION_SECRET is worse than no auth at all: HMAC with a
# known ("") key means anyone reading this open-source file can mint their
# own valid session tokens and skip the password check entirely, with no
# error or log line to reveal that's happening. An empty HUB_PASSWORD would
# do the same via the login endpoint. Both used to default to "" silently.
if not SESSION_SECRET:
    raise RuntimeError(
        "SESSION_SECRET env var is not set. Refusing to start: an empty "
        "HMAC secret would let anyone forge valid session tokens. Generate "
        "one with: python -c \"import secrets; print(secrets.token_hex(32))\""
    )
if not HUB_PASSWORD:
    raise RuntimeError(
        "HUB_ACCESS_PASSWORD env var is not set. Refusing to start: with no "
        "password configured, /auth/login would reject every login attempt "
        "anyway, so this is almost certainly a missed deploy config step."
    )

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
    # Added for the Market Map's radius click feature (1/3/10mi competitor
    # list + market share). Exhaustive spatial join, not size-filtered like
    # branch_target_competitors — this is the same source Power BI uses.
    "branch_competitors_10mi_v2":       "geo",
    # Rate Radar's latest-scrape-per-institution view — used by
    # rate-radar.html and the Hub's Rate Radar panel.
    "vw_rate_radar_latest":             "public",
    # Rate Radar's per-run trend view — powers the Hub's CD-rate sparkline.
    # Was missing entirely, so every sparkline request 403'd and silently
    # rendered empty.
    "vw_rate_radar_history":            "public",
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


# Rate Radar's "mark verified" action — a human confirmed this specific rate
# on the bank's own site. Narrowly scoped on purpose: unlike proxy_table,
# this can only touch rate_observations.manually_verified/verified_by/
# verified_at (never extraction_method/confidence, so the original automated
# classification is never overwritten and unverify is a clean, lossless
# toggle), and only for the one row identified by (bank_name, run_id,
# product_type). Anyone with a valid Hub session can call this today — there
# is no per-person role system, only the one shared Hub password.
VERIFY_PRODUCT_TYPES = {"checking", "savings", "high_yield_savings", "cd", "money_market"}

@secure_proxy_bp.route("/api/verify-rate", methods=["POST", "OPTIONS"])
@require_session
def verify_rate():
    if request.method == "OPTIONS":
        return _cors_headers(jsonify({}))

    body = request.get_json(force=True, silent=True) or {}
    bank_name    = (body.get("bank_name") or "").strip()
    run_id       = (body.get("run_id") or "").strip()
    product_type = (body.get("product_type") or "").strip()
    verified     = bool(body.get("verified"))
    verified_by  = (body.get("verified_by") or "").strip()

    if not bank_name or not run_id:
        return jsonify({"error": "bank_name and run_id are required"}), 400
    if product_type not in VERIFY_PRODUCT_TYPES:
        return jsonify({"error": f"product_type must be one of {sorted(VERIFY_PRODUCT_TYPES)}"}), 400
    if not verified_by:
        return jsonify({"error": "verified_by is required (who is marking this?)"}), 400

    row = {
        "manually_verified": verified,
        "verified_by": verified_by,
        "verified_at": datetime.now(timezone.utc).isoformat(),
    }
    url = (f"{SUPA_URL}/rest/v1/rate_observations"
           f"?bank_name=eq.{quote(bank_name)}&run_id=eq.{quote(run_id)}&product_type=eq.{quote(product_type)}")

    try:
        r = requests.patch(
            url,
            headers={
                "apikey": SUPA_SERVICE,
                "Authorization": f"Bearer {SUPA_SERVICE}",
                "Content-Type": "application/json",
                "Prefer": "return=representation",
            },
            json=row,
            timeout=20,
        )
        if r.status_code not in (200, 204):
            return jsonify({"error": r.text[:300]}), r.status_code
        updated = r.json() if r.text else []
        if not updated:
            return jsonify({"error": "no matching rate_observations row — check bank_name/run_id/product_type"}), 404
        return jsonify({"ok": True, "updated": len(updated)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500
