"""Direct Postgres connection for long-running pipeline procedures that must
bypass PostgREST entirely.

Why: Supabase's `authenticator` role (what PostgREST connects as for every
REST/RPC call) has an 8-second statement_timeout -- a platform-wide safety
default applying to every API call. Confirmed directly that this can't be
worked around from inside a called function (SET LOCAL statement_timeout
doesn't help -- the timeout is locked in when the OUTER statement begins,
before the function body even runs, so a mid-function override never takes
effect for that call). Every rebuild procedure built this week takes well
over 8 seconds (some run 10+ minutes), so calling them via
ingestion.supabase_client.call_rpc() (PostgREST) always fails once actually
exercised outside an interactive session.

A direct psycopg2 connection has no PostgREST/authenticator layer involved
at all, so statement_timeout can be set freely per-connection.

Requires SUPABASE_DB_URL -- the direct Postgres connection string (URI
format) from Supabase's dashboard: Project Settings -> Database ->
Connection string -> "Direct connection" (NOT the transaction pooler, which
doesn't support the multi-COMMIT stored procedures some of these call --
batched rebuilds COMMIT between chunks, which the pooler's transaction mode
can't handle correctly).
"""

import os

import psycopg2

DB_URL = os.environ.get("SUPABASE_DB_URL", "")

DEFAULT_TIMEOUT_MS = 30 * 60 * 1000  # 30 minutes


def call_procedure(sql, params=None, timeout_ms=DEFAULT_TIMEOUT_MS):
    """Runs one CALL/SELECT statement (a stored procedure) on a fresh direct
    connection, with statement_timeout raised for the session.

    autocommit=True is required, not optional: several of these procedures
    (the tiered/10mi competitor rebuilds) issue internal COMMIT statements
    to batch a long-running loop, and psycopg2 must not wrap the CALL in its
    own client-side transaction for that to work -- with autocommit off,
    Postgres rejects a COMMIT issued from inside a procedure that the client
    itself already wrapped in a transaction block.
    """
    if not DB_URL:
        raise RuntimeError(
            "SUPABASE_DB_URL env var is not set -- direct Postgres access "
            "requires it (Supabase dashboard: Project Settings -> Database "
            "-> Connection string -> Direct connection)."
        )
    conn = psycopg2.connect(DB_URL)
    conn.autocommit = True
    try:
        with conn.cursor() as cur:
            cur.execute(f"SET statement_timeout = {int(timeout_ms)}")
            cur.execute(sql, params)
    finally:
        conn.close()
