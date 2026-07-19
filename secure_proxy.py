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
# exactly what context-generator.html actually queries today.
# Each maps to the Postgres schema it actually lives in, so we can
# send the right Accept-Profile header — PostgREST defaults to
# 'public' when that header is missing, which is why dim_institutions
# (in 'ref') and the analytics-schema tables need it set explicitly.
ALLOWED_TABLES = {
    "dim_institutions":                 "ref",
    "branch_opportunity_base":          "analytics",
    "branch_target_competitors":        "analytics",
    "bank_financial_snapshot_latest":   "analytics",
    "vw_network_top_targets":           "public",
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


@secure_proxy_bp.route("/api/ai/briefing-note", methods=["POST", "OPTIONS"])
@require_session
def ai_briefing_note():
    """Replaces the direct-from-browser call to api.anthropic.com in
    the 'Tom's Briefing Note' feature. Browser sends the already-built
    context string; this endpoint holds the actual API key."""
    if request.method == "OPTIONS":
        return _cors_headers(jsonify({}))

    body = request.get_json(force=True, silent=True) or {}
    brief_context = (body.get("context") or "").strip()
    if not brief_context:
        return jsonify({"error": "context required"}), 400

    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "Content-Type": "application/json",
                "x-api-key": ANTH_KEY,
                "anthropic-version": "2023-06-01",
            },
            json={
                "model": "claude-sonnet-4-6",
                "max_tokens": 700,
                "system": (
                    "You are the BMAP Executive Strategist at Verlocity. Brief Tom "
                    "before he walks into a meeting. Sharp colleague tone. Return "
                    "ONLY a JSON object with keys: paragraph (string), strong "
                    "(array of 3 strings \"Strength — implication\"), pressure "
                    "(array of 3 strings \"Pressure — our angle\"), asymmetric "
                    "(string, 2-3 sentences with specific branch name or number "
                    "that would surprise a banker). No markdown, no explanation, "
                    "ONLY the JSON."
                ),
                "messages": [{"role": "user", "content": brief_context}],
            },
            timeout=30,
        )
        data = resp.json()
        txt = next((b["text"] for b in data.get("content", []) if b.get("type") == "text"), "{}")
        txt = txt.replace("```json", "").replace("```", "").strip()
        parsed = json.loads(txt)
        return jsonify(parsed)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
