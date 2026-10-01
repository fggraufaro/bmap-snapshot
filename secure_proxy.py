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

Optional (Phase 1 per-account auth, migrating off the shared passcode):
    SUPABASE_ANON_KEY      — Settings → API → anon/publishable key (safe server-
                              side; it's Supabase Auth's project identifier, not
                              a secret). Without it, /auth/login's email+password
                              path returns 503 but the shared passcode keeps
                              working unaffected — this is additive, not a
                              replacement, until every current user has a real
                              account.

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
from flask import Blueprint, request, jsonify, g

secure_proxy_bp = Blueprint("secure_proxy", __name__)

# ── Config ──────────────────────────────────────────────────────
SUPA_URL      = "https://tuiiywphoynbmkxpoyps.supabase.co"
SUPA_SERVICE  = os.environ.get("SUPABASE_SERVICE_KEY", "")
# Phase 1 per-account auth (email+password via Supabase Auth, alongside the
# existing shared passcode -- both work during the migration, see login()).
# The anon/publishable key is safe to hold server-side; it's the project
# identifier Supabase Auth's own token endpoint requires, not a secret.
# Optional, not fail-fast like SESSION_SECRET/HUB_PASSWORD below: missing
# it should only break the new email-login path, not take down the
# passcode login everyone currently depends on.
SUPA_ANON_KEY = os.environ.get("SUPABASE_ANON_KEY", "")
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
    # a62: Growth Map's new-signal sandbox — CFPB week-over-week complaint
    # trend per institution, first signal tested here before any polished
    # Opportunity View build.
    "vw_cfpb_complaints_wow":           "public",
    # Growth Map's YoY delta overlay — 2026 vs 2025 opportunity_score per
    # branch, computed server-side (joins the current table against the
    # 2025 archive) so the browser never has to fetch and diff two years
    # of data itself.
    "vw_branch_opportunity_yoy_delta":  "public",
}

# Phase 2: per-table scoping for role=bank_user. Verified against the real
# schema (not assumed) -- the institution-key space is NOT consistent across
# these tables, same landmine CLAUDE.md already documents elsewhere
# (RSSDID vs CERT, CITYBR vs CITY). Confirmed live: inst_key == 'bank_' +
# rssdid (e.g. 'bank_1000052' / 1000052), so tables keyed on a raw RSSD
# integer are reachable from a profile's inst_key by stripping the prefix.
#   "direct" -- column holds inst_key text directly, equality filter.
#   "rssd"   -- column holds the raw RSSD integer/text, derived from inst_key.
# "my_inst_key" tables (branch_target_competitors, vw_network_top_targets)
# are competitive-intelligence views: scoping by the *_inst_key_ side is
# correct because the point of those views is showing a bank's own
# branches against surrounding competitors -- the competitor/target side
# legitimately contains other institutions by design.
TABLE_SCOPE = {
    "dim_institutions":                ("inst_key", "direct"),
    "bank_financial_snapshot_latest":  ("inst_key", "direct"),
    "branch_opportunity_base":         ("inst_key", "direct"),
    "vw_branch_opportunity_cbsa":      ("inst_key", "direct"),
    "vw_branch_opportunity_yoy_delta": ("inst_key", "direct"),
    "vw_cfpb_complaints_wow":          ("inst_key", "direct"),
    "vw_prospecting_score":            ("inst_key", "direct"),
    "branch_target_competitors":       ("my_inst_key", "direct"),
    "vw_network_top_targets":          ("my_inst_key", "direct"),
    "branch_competitors_10mi_v2":      ("my_bank_id", "rssd"),
    "vw_rate_radar_latest":            ("rssdid", "rssd"),
    "vw_rate_radar_history":           ("rssdid", "rssd"),
    "bank_website":                    ("FED_RSSD", "rssd"),
}
# Tables with no institution dimension at all (verified: no inst_key/rssd/
# cert/bank_id column exists) -- geography/demographic reference data,
# identical for every role. Anything in ALLOWED_TABLES that is in neither
# this set nor TABLE_SCOPE is blocked for bank_user (fail closed), so a
# newly added table can't go unscoped just because no one classified it yet.
GLOBAL_TABLES = {"uszips", "vw_zip_persona"}

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
# role/inst_key are baked into the signed payload (not re-looked-up per
# request) so proxy_table/proxy_rpc can trust g.role/g.inst_key the same
# way the rest of this file trusts a verified signature -- tampering with
# either field invalidates the signature. Passcode logins (no email) get
# role="admin" with no inst_key, matching their historical full access.
def _make_token(subject: str = "hub", role: str = "admin", inst_key: str = "") -> str:
    exp = int(time.time()) + SESSION_TTL_SECONDS
    payload = f"{subject}:{role}:{inst_key}:{exp}"
    sig = hmac.new(SESSION_SECRET.encode(), payload.encode(), hashlib.sha256).hexdigest()
    raw = f"{payload}:{sig}"
    return base64.urlsafe_b64encode(raw.encode()).decode()


def _decode_token(token: str):
    try:
        raw = base64.urlsafe_b64decode(token.encode()).decode()
        subject, role, inst_key, exp, sig = raw.split(":")
        payload = f"{subject}:{role}:{inst_key}:{exp}"
        expected = hmac.new(SESSION_SECRET.encode(), payload.encode(), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(sig, expected):
            return None
        if int(exp) < time.time():
            return None
        return {"subject": subject, "role": role, "inst_key": inst_key}
    except Exception:
        return None


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
        identity = _decode_token(token) if token else None
        if not identity:
            return jsonify({"error": "unauthorized"}), 401
        g.role = identity["role"]
        g.inst_key = identity["inst_key"]
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
    email = (body.get("email") or "").strip()
    password = (body.get("password") or "").strip()

    attempts.append(now)
    _login_attempts[ip] = attempts

    # Per-account path: email present means this is a real Supabase Auth
    # login, not the shared passcode. Runs entirely server-side -- the
    # browser never receives or holds a Supabase key, same property the
    # rest of this file maintains for the Anthropic key and the service
    # role. Both this and the shared-passcode path below stay live at the
    # same time during the migration; see Phase 1 notes in the roadmap.
    if email:
        if not SUPA_ANON_KEY:
            return jsonify({"error": "email login not configured yet"}), 503
        try:
            r = requests.post(
                f"{SUPA_URL}/auth/v1/token?grant_type=password",
                headers={"apikey": SUPA_ANON_KEY, "Content-Type": "application/json"},
                json={"email": email, "password": password},
                timeout=10,
            )
        except Exception:
            return jsonify({"error": "login service unavailable"}), 503
        if r.status_code != 200:
            return jsonify({"error": "incorrect email or password"}), 401
        user_id = (r.json().get("user") or {}).get("id")
        if not user_id:
            return jsonify({"error": "incorrect email or password"}), 401

        # Look up role/inst_key -- service role, bypasses RLS. This lookup
        # *is* the trust boundary (same pattern as every other Supabase
        # call in this file): a valid Supabase session alone doesn't grant
        # Hub access, a matching profiles row does.
        try:
            prof_r = requests.get(
                f"{SUPA_URL}/rest/v1/profiles?id=eq.{user_id}&select=role,inst_key,email",
                headers={"apikey": SUPA_SERVICE, "Authorization": f"Bearer {SUPA_SERVICE}"},
                timeout=10,
            )
            profiles = prof_r.json() if prof_r.ok else []
        except Exception:
            profiles = []
        if not profiles:
            return jsonify({"error": "no profile configured for this account — contact an admin"}), 403

        profile = profiles[0]
        role = profile.get("role") or "bank_user"
        inst_key = profile.get("inst_key") or ""
        if role == "bank_user" and not inst_key:
            return jsonify({"error": "account has no institution assigned — contact an admin"}), 403

        _login_attempts[ip] = []  # reset on success
        token = _make_token(subject=user_id, role=role, inst_key=inst_key)
        return jsonify({
            "token": token,
            "expires_in": SESSION_TTL_SECONDS,
            "role": role,
            "inst_key": inst_key,
            "email": profile.get("email"),
        })

    # Existing path: the shared passcode. Unchanged -- stays working for
    # anyone not yet migrated to a real account.
    if not HUB_PASSWORD or not hmac.compare_digest(password, HUB_PASSWORD):
        return jsonify({"error": "incorrect password"}), 401

    _login_attempts[ip] = []  # reset on success
    token = _make_token()
    return jsonify({"token": token, "expires_in": SESSION_TTL_SECONDS})


@secure_proxy_bp.route("/auth/set-password", methods=["POST", "OPTIONS"])
def set_password():
    # Called right after someone clicks a real invite/recovery email link.
    # Supabase already verified that click server-side and handed the
    # browser a short-lived access_token in the URL -- this endpoint uses
    # that token to actually set the password, still without the browser
    # ever holding a Supabase key itself (same property as /auth/login).
    if request.method == "OPTIONS":
        return _cors_headers(jsonify({}))

    if not SUPA_ANON_KEY:
        return jsonify({"error": "not configured yet"}), 503

    body = request.get_json(force=True, silent=True) or {}
    access_token = (body.get("access_token") or "").strip()
    new_password = body.get("new_password") or ""

    if not access_token or len(new_password) < 8:
        return jsonify({"error": "invalid request"}), 400

    try:
        r = requests.put(
            f"{SUPA_URL}/auth/v1/user",
            headers={
                "apikey": SUPA_ANON_KEY,
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            },
            json={"password": new_password},
            timeout=10,
        )
    except Exception:
        return jsonify({"error": "service unavailable"}), 503

    if r.status_code != 200:
        return jsonify({"error": "could not set password — the link may have expired, request a new one"}), 400

    return jsonify({"ok": True})


@secure_proxy_bp.route("/api/<table>", methods=["GET", "OPTIONS"])
@require_session
def proxy_table(table):
    if request.method == "OPTIONS":
        return _cors_headers(jsonify({}))

    schema = ALLOWED_TABLES.get(table)
    if schema is None:
        return jsonify({"error": f"table '{table}' is not exposed via the proxy"}), 403

    # Forward the querystring as-is (select=, filters, order, limit —
    # these are the same params the Hub already builds client-side) —
    # then, for a scoped bank_user, append a server-side institution
    # filter PostgREST ANDs against it. The client's own querystring is
    # never trusted for the security boundary, only for the forced filter
    # we add here, so a bank_user can't widen their own access by
    # omitting or editing filters client-side.
    qs = request.query_string.decode()

    if g.role != "admin":
        if table in GLOBAL_TABLES:
            pass  # no institution dimension — same for every role
        elif table in TABLE_SCOPE:
            column, mode = TABLE_SCOPE[table]
            if not g.inst_key:
                return jsonify({"error": "no institution assigned to this account"}), 403
            if mode == "rssd":
                value = g.inst_key.split("bank_", 1)[-1]
            else:
                value = g.inst_key
            qs = f"{qs}&{column}=eq.{quote(value)}" if qs else f"{column}=eq.{quote(value)}"
        else:
            # Not yet classified for scoping — fail closed rather than
            # silently serving an unscoped table to a non-admin.
            return jsonify({"error": f"table '{table}' is not yet available for this account type"}), 403

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

    # These four are Growth Map's radius-click spatial queries, not yet
    # reviewed for institution scoping and not yet rolled out to
    # bank_user accounts anyway (Growth Map still only has the shared-
    # passcode gate). Fail closed until that gate gets the same role
    # check this file now has for /api/<table>.
    if g.role != "admin":
        return jsonify({"error": "not available for this account type"}), 403

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


# Rate Radar's "mark verified" / "correct this value" actions. Narrowly
# scoped on purpose: unlike proxy_table, these can only touch
# rate_observations.apy/manually_verified/verified_by/verified_at (never
# confidence directly) for the one row identified by (bank_name, run_id,
# product_type). Anyone with a valid Hub session can call this today —
# there is no per-person role system, only the one shared Hub password.
VERIFY_PRODUCT_TYPES = {"checking", "savings", "high_yield_savings", "cd", "money_market"}
VERIFY_MAX_AGE_DAYS = 7  # only recent runs can be verified/edited — a stale
                         # row should get a fresh crawl, not a manual patch
                         # that a real re-scrape would just overwrite anyway.

def _run_age_days(run_id):
    """Days since the run started, or None if the run can't be found."""
    try:
        r = requests.get(
            f"{SUPA_URL}/rest/v1/rate_radar_runs?run_id=eq.{quote(run_id)}&select=started_at",
            headers={"apikey": SUPA_SERVICE, "Authorization": f"Bearer {SUPA_SERVICE}"},
            timeout=10,
        )
        rows = r.json()
        if not rows or not rows[0].get("started_at"):
            return None
        started = datetime.fromisoformat(rows[0]["started_at"].replace("Z", "+00:00"))
        return (datetime.now(timezone.utc) - started).days
    except Exception:
        return None


def _check_verify_request(body):
    """Shared validation for verify-rate and edit-rate. Returns
    (bank_name, run_id, product_type, verified_by, error_response)."""
    # Rate Radar curation (marking/correcting scraped rates) touches shared
    # reference data across every institution, not a per-bank self-service
    # action -- admin only, same fail-closed default as proxy_table/proxy_rpc.
    if g.role != "admin":
        return None, None, None, None, (jsonify({"error": "not available for this account type"}), 403)

    bank_name    = (body.get("bank_name") or "").strip()
    run_id       = (body.get("run_id") or "").strip()
    product_type = (body.get("product_type") or "").strip()
    verified_by  = (body.get("verified_by") or "").strip()

    if not bank_name or not run_id:
        return None, None, None, None, (jsonify({"error": "bank_name and run_id are required"}), 400)
    if product_type not in VERIFY_PRODUCT_TYPES:
        return None, None, None, None, (jsonify({"error": f"product_type must be one of {sorted(VERIFY_PRODUCT_TYPES)}"}), 400)
    if not verified_by:
        return None, None, None, None, (jsonify({"error": "verified_by is required (who is marking this?)"}), 400)

    age = _run_age_days(run_id)
    if age is None:
        return None, None, None, None, (jsonify({"error": "couldn't find that run — refresh the page and try again"}), 404)
    if age > VERIFY_MAX_AGE_DAYS:
        return None, None, None, None, (jsonify({
            "error": f"this data is {age} days old — verifying/editing is only allowed within "
                     f"{VERIFY_MAX_AGE_DAYS} days of the crawl. Trigger a fresh crawl instead."
        }), 400)

    return bank_name, run_id, product_type, verified_by, None


def _patch_rate_observation(bank_name, run_id, product_type, row):
    url = (f"{SUPA_URL}/rest/v1/rate_observations"
           f"?bank_name=eq.{quote(bank_name)}&run_id=eq.{quote(run_id)}&product_type=eq.{quote(product_type)}")
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
    return r


@secure_proxy_bp.route("/api/verify-rate", methods=["POST", "OPTIONS"])
@require_session
def verify_rate():
    if request.method == "OPTIONS":
        return _cors_headers(jsonify({}))

    body = request.get_json(force=True, silent=True) or {}
    bank_name, run_id, product_type, verified_by, err = _check_verify_request(body)
    if err:
        return err
    verified = bool(body.get("verified"))

    row = {
        "manually_verified": verified,
        "verified_by": verified_by,
        "verified_at": datetime.now(timezone.utc).isoformat(),
    }
    try:
        r = _patch_rate_observation(bank_name, run_id, product_type, row)
        if r.status_code not in (200, 204):
            return jsonify({"error": r.text[:300]}), r.status_code
        updated = r.json() if r.text else []
        if not updated:
            return jsonify({"error": "no matching rate_observations row — check bank_name/run_id/product_type"}), 404
        return jsonify({"ok": True, "updated": len(updated)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# "Correct this value" — unlike a plain verify (confirming the crawler's
# number was already right), this overwrites apy with what a human typed in,
# so it's honestly labeled extraction_method='manual' rather than keeping
# whatever the automated crawl originally guessed for a value that's since
# been overwritten. Always implies manually_verified=true.
@secure_proxy_bp.route("/api/edit-rate", methods=["POST", "OPTIONS"])
@require_session
def edit_rate():
    if request.method == "OPTIONS":
        return _cors_headers(jsonify({}))

    body = request.get_json(force=True, silent=True) or {}
    bank_name, run_id, product_type, verified_by, err = _check_verify_request(body)
    if err:
        return err

    try:
        apy = float(body.get("apy"))
    except (TypeError, ValueError):
        return jsonify({"error": "apy must be a number"}), 400
    if not (0 <= apy <= 15):
        return jsonify({"error": "apy must be between 0 and 15"}), 400

    row = {
        "apy": apy,
        "extraction_method": "manual",
        "confidence": "high",
        "manually_verified": True,
        "verified_by": verified_by,
        "verified_at": datetime.now(timezone.utc).isoformat(),
    }
    try:
        r = _patch_rate_observation(bank_name, run_id, product_type, row)
        if r.status_code not in (200, 204):
            return jsonify({"error": r.text[:300]}), r.status_code
        updated = r.json() if r.text else []
        if not updated:
            return jsonify({"error": "no matching rate_observations row — check bank_name/run_id/product_type"}), 404
        return jsonify({"ok": True, "updated": len(updated)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500
