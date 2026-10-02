"""Shared PostgREST helper for ingestion scripts.

Mirrors the request pattern already used in bmap_snapshot.py (supabase() /
supabase_insert()) rather than introducing a new client library or
convention. Non-public schemas need Accept-Profile / Content-Profile
headers — see CLAUDE.md's "Schema routing" gotcha.
"""

import os
import time

import requests

SUPA_URL = "https://tuiiywphoynbmkxpoyps.supabase.co"
SUPA_KEY = os.environ.get("SUPABASE_SERVICE_KEY", "")

if not SUPA_KEY:
    print("WARNING: SUPABASE_SERVICE_KEY is not set - Supabase calls will fail with 401.")

# Transient connection drops (SSL EOF, RemoteDisconnected, etc.) on large
# batched uploads -- seen 3x in one session pulling multi-hundred-thousand
# to multi-million row ingestions (FFIEC UBPR peer-stats/rank) -- not a
# logic bug, just a flaky connection on big transfers. Retried with backoff
# rather than requiring a manual re-run each time; upsert is idempotent so a
# retried batch is always safe.
_RETRYABLE = (requests.exceptions.ConnectionError, requests.exceptions.ChunkedEncodingError,
              requests.exceptions.Timeout)


def _post_with_retry(url, headers, json_body, timeout, max_attempts=4):
    for attempt in range(1, max_attempts + 1):
        try:
            r = requests.post(url, headers=headers, json=json_body, timeout=timeout)
            r.raise_for_status()
            return r
        except _RETRYABLE as e:
            if attempt == max_attempts:
                raise
            wait = 2 ** attempt  # 2, 4, 8s
            print(f"  transient error ({e!r}), retrying in {wait}s (attempt {attempt}/{max_attempts})...")
            time.sleep(wait)


def _headers(schema, write=False):
    h = {"apikey": SUPA_KEY, "Authorization": f"Bearer {SUPA_KEY}"}
    if schema != "public":
        h["Content-Profile" if write else "Accept-Profile"] = schema
    return h


def get(table, schema="public", params=""):
    url = f"{SUPA_URL}/rest/v1/{table}"
    if params:
        url += f"?{params}"
    r = requests.get(url, headers=_headers(schema), timeout=30)
    r.raise_for_status()
    return r.json()


def upsert(table, rows, on_conflict, schema="public", batch_size=5000):
    """Batched POST with Prefer: resolution=merge-duplicates (upsert).

    Returns the total number of rows sent (PostgREST doesn't echo row
    counts back with Prefer: return=minimal, so this is the send count,
    not a confirmed write count).
    """
    if not rows:
        return 0
    url = f"{SUPA_URL}/rest/v1/{table}?on_conflict={on_conflict}"
    headers = _headers(schema, write=True)
    headers["Content-Type"] = "application/json"
    headers["Prefer"] = "resolution=merge-duplicates,return=minimal"

    sent = 0
    for i in range(0, len(rows), batch_size):
        batch = rows[i:i + batch_size]
        _post_with_retry(url, headers, batch, timeout=120)
        sent += len(batch)
    return sent


def call_rpc(fn, args=None, schema="public", timeout=1800):
    """POST to /rest/v1/rpc/<fn> - for the refresh_*/archive_* stored
    procedures (refresh_bmap_after_upload, refresh_branch_opportunity_base,
    refresh_bmap_scores, refresh_branches_master_v2, archive_bmap_year_snapshot,
    ...) that ingestion scripts and the pipeline command center trigger.

    `args` is a dict of the function's named parameters (e.g. {"p_year": 2026})
    -- PostgREST maps these directly to the Postgres function's argument names.
    Default timeout is 30 minutes: these can be genuine multi-minute rebuilds
    (the tiered/10mi competitor systems took 10+ minutes end to end), and the
    call is expected to run from a background job, not an interactive request,
    so there's no reason to cut it short the way a page-load request would need
    to be."""
    url = f"{SUPA_URL}/rest/v1/rpc/{fn}"
    headers = _headers(schema, write=True)
    headers["Content-Type"] = "application/json"
    r = requests.post(url, headers=headers, json=(args or {}), timeout=timeout)
    r.raise_for_status()
    return r.json() if r.content else None
