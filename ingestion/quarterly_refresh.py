"""New-quarter refresh steps (a84): the UBPR analytics layer ("step A") and the
bank financial snapshot ("step B"), each as a dry run and a swap.

Spec: Session 3's a84 pipeline spec. Both are stage -> validate -> back up ->
swap, because the slow work (the view / the DISTINCT ON over millions of rows)
happens in a staging table with no locks on production, and the swap itself is
a few seconds inside ONE transaction.

DRY RUN  (snapshot_dry_run / ubpr_layer_dry_run): precheck, stage, validate,
         report, drop staging. Never touches a production table.
SWAP     (snapshot_swap / ubpr_layer_swap): everything above, then back up and
         TRUNCATE + INSERT the production tables. DESTRUCTIVE, so it refuses
         to run unless the env var ALLOW_ANALYTICS_SWAP=yes is set on the
         service. It is NOT set: the a90 hold is in force and the Coordinator
         must sign off before the first real run. Setting it is the unlock.

Never calls refresh_bmap_after_upload. Never touches branch_opportunity_base or
branch_target_competitors (a90 hold).

Design choices worth knowing:
  - Staging tables are named stg_pipeline_* (not the spec's stg_bank_financial_
    snapshot) so a leftover staging table from another session is never
    dropped by this code.
  - INSERTs use explicit column lists. analytics.bank_financial_snapshot_latest
    has two more columns than the view (marketing_expenses, marketing_ratio);
    a positional SELECT * would silently NULL them, so the swap aborts if they
    ever hold data, and otherwise leaves them NULL on purpose.
  - The network_top_targets update runs INSIDE the swap transaction, so a
    failure rolls back the snapshot too (the spec ran it as a separate step).
    It only writes rows whose values actually change.
  - After the inserts, counts are re-checked inside the transaction and any
    mismatch raises, which rolls everything back.
  - Backup tables are suffixed with date AND time so a same-day retry can't
    collide.

ROLLBACK (after a swap), in one transaction, from the backup.* tables named in
the job's log:
    TRUNCATE analytics.bank_financial_snapshot;
    INSERT INTO analytics.bank_financial_snapshot SELECT * FROM backup.bank_financial_snapshot_<ts>;
    TRUNCATE analytics.bank_financial_snapshot_latest;
    INSERT INTO analytics.bank_financial_snapshot_latest SELECT * FROM backup.bank_financial_snapshot_latest_<ts>;
    UPDATE public.network_top_targets t SET target_roa=b.target_roa, target_efficiency_ratio=b.target_efficiency_ratio,
      target_noncurrent_pct=b.target_noncurrent_pct, target_cof_pct=b.target_cof_pct,
      target_dep_yoy_financial=b.target_dep_yoy_financial
      FROM backup.network_top_targets_<ts> b WHERE b.my_inst_key=t.my_inst_key AND b.target_inst_key=t.target_inst_key;
(UBPR layer: TRUNCATE + INSERT each analytics.ubpr_* table from backup.ubpr_*_<ts>.)

CLI:
  python -m ingestion.quarterly_refresh snapshot --dry-run
  python -m ingestion.quarterly_refresh ubpr --dry-run
  python -m ingestion.quarterly_refresh snapshot|ubpr          # swap; needs ALLOW_ANALYTICS_SWAP=yes
"""

import datetime
import os
import sys
from contextlib import contextmanager

import psycopg2
from psycopg2 import sql

from ingestion.ffiec_bulk_portal import CORE_METRIC_SUFFIXES

GUARD_ENV = "ALLOW_ANALYTICS_SWAP"
LOCK_KEY = 84_850_001           # pg advisory lock: one refresh at a time
STAGE_TIMEOUT = "3000s"
N_CODES = len(CORE_METRIC_SUFFIXES)   # 131: every period must carry all of them

PROD = {
    "snapshot": ("analytics", "bank_financial_snapshot"),
    "latest": ("analytics", "bank_financial_snapshot_latest"),
    "ntt": ("public", "network_top_targets"),
    "peer_stats": ("analytics", "ubpr_peer_stats_clean"),
    "rank": ("analytics", "ubpr_rank_clean"),
    "bank_pg": ("analytics", "ubpr_bank_peer_group"),
    "coverage": ("analytics", "ubpr_rank_coverage"),
}
STG = {  # this code's own staging tables, safe to drop
    "snapshot": ("analytics", "stg_pipeline_bank_financial_snapshot"),
    "peer_stats": ("analytics", "stg_pipeline_ubpr_peer_stats_clean"),
    "rank": ("analytics", "stg_pipeline_ubpr_rank_clean"),
    "bank_pg": ("analytics", "stg_pipeline_ubpr_bank_peer_group"),
    "coverage": ("analytics", "stg_pipeline_ubpr_rank_coverage"),
}
TIER_GROUPS = ("1", "2", "3", "4", "5", "6", "7", "8", "101", "102", "103", "104",
               "201", "202", "203", "301", "401")


def _log(msg=""):
    print(msg, flush=True)


def _ident(t):
    return sql.Identifier(*t)


def _check_guard():
    if os.environ.get(GUARD_ENV) != "yes":
        raise RuntimeError(
            f"Swap is locked: {GUARD_ENV} is not set to 'yes' on this service. The a90 hold is in force; the "
            "Coordinator must sign off before the first real run. Use the dry-run step to test safely."
        )


@contextmanager
def _session():
    """One direct connection (autocommit; statements that need a transaction open
    their own) holding the advisory lock for the duration."""
    url = os.environ.get("SUPABASE_DB_URL", "")
    if not url:
        raise RuntimeError("SUPABASE_DB_URL is not set -- direct Postgres access is required.")
    conn = psycopg2.connect(url)
    conn.autocommit = True
    try:
        with conn.cursor() as cur:
            cur.execute(f"SET statement_timeout = '{STAGE_TIMEOUT}'")
            cur.execute("SET work_mem = '256MB'")
            cur.execute("SELECT pg_try_advisory_lock(%s)", (LOCK_KEY,))
            if not cur.fetchone()[0]:
                raise RuntimeError("Another quarterly refresh is already running (advisory lock held); not starting.")
        yield conn
    finally:
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT pg_advisory_unlock(%s)", (LOCK_KEY,))
        finally:
            conn.close()


def _q(cur, query, params=None):
    cur.execute(query, params)
    return cur.fetchall()


def _scalar(cur, query, params=None):
    cur.execute(query, params)
    return cur.fetchone()[0]


def _cols(cur, table):
    cur.execute("select column_name from information_schema.columns where table_schema=%s and table_name=%s "
                "order by ordinal_position", table)
    return [r[0] for r in cur.fetchall()]


def _drop(cur, table):
    cur.execute(sql.SQL("DROP TABLE IF EXISTS {}").format(_ident(table)))


def _lock_down(cur, table):
    cur.execute(sql.SQL("ALTER TABLE {} ENABLE ROW LEVEL SECURITY").format(_ident(table)))
    cur.execute(sql.SQL("REVOKE ALL ON {} FROM anon, authenticated").format(_ident(table)))


class Checks:
    def __init__(self):
        self.rows = []

    def add(self, name, ok, detail):
        self.rows.append((name, bool(ok), detail))
        _log(f"  [{'PASS' if ok else 'FAIL'}] {name}: {detail}")

    def info(self, name, detail):
        _log(f"  [info] {name}: {detail}")

    @property
    def ok(self):
        return all(ok for _, ok, _ in self.rows)

    def raise_if_failed(self):
        bad = [f"{n} ({d})" for n, ok, d in self.rows if not ok]
        if bad:
            raise RuntimeError("Validation failed, nothing was swapped: " + "; ".join(bad))


# --------------------------------------------------------------------------
# Step B: financial snapshot
# --------------------------------------------------------------------------

def _snapshot_precheck(cur):
    """New quarter must be present in all three raw tables, with sane coverage."""
    checks = Checks()
    new = _scalar(cur, 'select max(period) from raw."raw_schedule_RI"')
    checks.info("newest quarter in raw_schedule_RI", new)
    for t in ("raw_schedule_RC", "raw_UBPR"):
        n = _scalar(cur, sql.SQL('select count(*) from raw.{} where period=%s').format(sql.Identifier(t)), (new,))
        checks.add(f"{t} has {new}", n > 0, f"{n} rows")
    ri_n = _scalar(cur, 'select count(*) from raw."raw_schedule_RI" where period=%s', (new,))
    rc_n = _scalar(cur, 'select count(*) from raw."raw_schedule_RC" where period=%s', (new,))
    prior = _scalar(cur, 'select max(period) from raw."raw_schedule_RI" where period < %s', (new,))
    prior_n = _scalar(cur, 'select count(*) from raw."raw_schedule_RI" where period=%s', (prior,)) if prior else None
    if prior_n:
        d = abs(ri_n - prior_n) / prior_n
        checks.add("bank count within 3% of prior quarter", d <= 0.03, f"{ri_n} vs {prior_n} ({prior}): {d * 100:.1f}%")
    for tbl, col in (("raw_schedule_RI", "RIAD4073"), ("raw_schedule_RI", "RIAD4074"),
                     ("raw_schedule_RI", "RIAD4340"), ("raw_schedule_RC", "RCON2200")):
        pop = _scalar(cur, sql.SQL('select count({c})::float / nullif(count(*),0) from raw.{t} where period=%s')
                      .format(c=sql.Identifier(col), t=sql.Identifier(tbl)), (new,))
        checks.add(f"{col} populated >= 99%", pop is not None and pop >= 0.99, f"{(pop or 0) * 100:.1f}%")
    checks.raise_if_failed()
    return new, rc_n


def _stage_snapshot(cur):
    _drop(cur, STG["snapshot"])
    _log("  staging from public.vw_bank_financial_snapshot (takes minutes) ...")
    t0 = datetime.datetime.now()
    cur.execute(sql.SQL("CREATE TABLE {} AS SELECT * FROM public.vw_bank_financial_snapshot").format(_ident(STG["snapshot"])))
    _lock_down(cur, STG["snapshot"])
    n = _scalar(cur, sql.SQL("select count(*) from {}").format(_ident(STG["snapshot"])))
    _log(f"  staged {n} rows in {(datetime.datetime.now() - t0).total_seconds():.0f}s")


def _validate_snapshot(cur, new_period, rc_banks, targets=PROD, stg=STG["snapshot"]):
    c = Checks()
    S, L = _ident(stg), _ident(targets["snapshot"])
    stg_rows = _scalar(cur, sql.SQL("select count(*) from {}").format(S))
    live_rows = _scalar(cur, sql.SQL("select count(*) from {}").format(L))
    c.add("staging rows >= live snapshot rows", stg_rows >= live_rows, f"{stg_rows} vs {live_rows}")
    dups = _scalar(cur, sql.SQL("select count(*) - count(distinct (inst_key, period)) from {}").format(S))
    c.add("no duplicate (inst_key, period) in staging", dups == 0, f"{dups} duplicates")
    scols, lcols = set(_cols(cur, stg)), set(_cols(cur, targets["snapshot"]))
    c.add("staging columns == live snapshot columns", scols == lcols,
          f"only in staging {sorted(scols - lcols)}, only in live {sorted(lcols - scols)}")

    newest = _scalar(cur, sql.SQL("select max(period) from {}").format(S))
    c.add("newest staged period == newest raw quarter", newest == new_period, f"{newest} vs {new_period}")
    banks = _scalar(cur, sql.SQL("select count(*) from {} where institution_type='bank' and period=%s").format(S), (new_period,))
    d = abs(banks - rc_banks) / rc_banks if rc_banks else 1
    c.add("banks at newest period within 3% of RC bank count", d <= 0.03, f"{banks} vs {rc_banks} ({d * 100:.1f}%)")
    nullshare = _scalar(cur, sql.SQL("select count(*) filter (where cost_of_funds_pct is null)::float / nullif(count(*),0) "
                                     "from {} where institution_type='bank' and period=%s").format(S), (new_period,))
    c.add("NULL cost_of_funds_pct share (banks, newest period) <= 3%", nullshare is not None and nullshare <= 0.03,
          f"{(nullshare or 0) * 100:.1f}%")
    med = _scalar(cur, sql.SQL("select percentile_cont(0.5) within group (order by cost_of_funds_pct) from {} "
                               "where institution_type='bank' and period=%s").format(S), (new_period,))
    c.add("median bank cost_of_funds_pct between 0.5 and 4.0", med is not None and 0.5 <= float(med) <= 4.0,
          f"{float(med):.2f}" if med is not None else "NULL")

    same = _q(cur, sql.SQL("""
        select count(*),
          count(*) filter (where s.total_assets is not distinct from l.total_assets
                             and s.roa is not distinct from l.roa
                             and s.efficiency_ratio is not distinct from l.efficiency_ratio
                             and s.loans_to_deposits_pct is not distinct from l.loans_to_deposits_pct),
          count(*) filter (where s.cost_of_funds_pct is not distinct from l.cost_of_funds_pct
                             and s.save_interest_pct is not distinct from l.save_interest_pct
                             and s.cd_interest_pct is not distinct from l.cd_interest_pct
                             and s.noninterest_income_pct is not distinct from l.noninterest_income_pct)
        from {} s join {} l using (inst_key, period)""").format(S, L))[0]
    pairs, nc_same, cost_same = same
    share = nc_same / pairs if pairs else 0
    c.add("non-cost columns identical on (inst_key, period) pairs in both >= 99.9%", share >= 0.999,
          f"{nc_same}/{pairs} ({share * 100:.2f}%)")
    cost_share = cost_same / pairs if pairs else 0
    no_new_data = _scalar(cur, sql.SQL("select (select max(period) from {}) = (select max(period) from {})").format(S, L))
    if no_new_data:
        # Same periods as live: the staged output must reproduce live, cost columns included.
        c.add("no new quarter: cost columns reproduce live snapshot >= 99.9%", cost_share >= 0.999,
              f"{cost_same}/{pairs} ({cost_share * 100:.2f}%)")
    else:
        c.info("cost columns identical to live on shared pairs (may differ after a view change)",
               f"{cost_same}/{pairs} ({cost_share * 100:.2f}%)")

    extra = [col for col in _cols(cur, targets["latest"]) if col not in scols]
    for col in extra:
        n = _scalar(cur, sql.SQL("select count({}) from {}").format(sql.Identifier(col), _ident(targets["latest"])))
        c.add(f"live latest.{col} (not in the view) holds no data that the swap would wipe", n == 0, f"{n} non-null")
    _log("  per-period row counts (staging vs live):")
    for p, sn, ln in _q(cur, sql.SQL("select p, (select count(*) from {S} where period=p), (select count(*) from {L} where period=p) "
                                     "from (select distinct period p from {S} union select distinct period from {L}) x "
                                     "order by p desc limit 6").format(S=S, L=L)):
        _log(f"    {p}: {sn} vs {ln}")
    return c


def _ntt_changes(cur, targets=PROD, stg=STG["snapshot"]):
    return _scalar(cur, sql.SQL("""
        with f as (select distinct on (inst_key) inst_key, roa, efficiency_ratio, noncurrent_assets_pct, cost_of_funds_pct, dep_yoy_pct
                   from {S} order by inst_key, period desc)
        select count(*) from {N} t join f on f.inst_key = t.target_inst_key
        where (t.target_roa, t.target_efficiency_ratio, t.target_noncurrent_pct, t.target_cof_pct, t.target_dep_yoy_financial)
              is distinct from (f.roa, f.efficiency_ratio, f.noncurrent_assets_pct, f.cost_of_funds_pct, f.dep_yoy_pct)
        """).format(S=_ident(stg), N=_ident(targets["ntt"])))


def snapshot_dry_run():
    _log("Snapshot refresh DRY RUN: stage + validate only, no production table is touched.")
    with _session() as conn, conn.cursor() as cur:
        try:
            new_period, rc_banks = _snapshot_precheck(cur)
            _stage_snapshot(cur)
            checks = _validate_snapshot(cur, new_period, rc_banks)
            _log(f"  network_top_targets rows a swap would update: {_ntt_changes(cur)}")
            checks.raise_if_failed()
            _log("DRY RUN PASSED: a real swap would validate cleanly on today's data.")
        finally:
            _drop(cur, STG["snapshot"])


def _backup(cur, src, schema, base, suffix):
    dst = (schema, f"{base}_{suffix}")
    cur.execute(sql.SQL("CREATE TABLE {} AS SELECT * FROM {}").format(_ident(dst), _ident(src)))
    _lock_down(cur, dst)
    n = _scalar(cur, sql.SQL("select count(*) from {}").format(_ident(dst)))
    src_n = _scalar(cur, sql.SQL("select count(*) from {}").format(_ident(src)))
    if n != src_n:
        raise RuntimeError(f"backup {dst[0]}.{dst[1]} has {n} rows but source has {src_n}; aborting before any swap")
    _log(f"  backed up {src[0]}.{src[1]} -> {dst[0]}.{dst[1]} ({n} rows)")
    return dst


def _swap_snapshot(conn, stg, targets, backup_schema, suffix):
    """Back up, then swap in ONE transaction. `targets`/`backup_schema` are parameters
    only so the swap SQL can be exercised against scratch copies (see
    analysis/a84_swap_selftest.py); real runs use PROD / 'backup'."""
    with conn.cursor() as cur:
        cols = _cols(cur, targets["snapshot"])
        col_sql = sql.SQL(", ").join(sql.Identifier(c) for c in cols)
        for key, base in (("snapshot", "bank_financial_snapshot"), ("latest", "bank_financial_snapshot_latest"),
                          ("ntt", "network_top_targets")):
            _backup(cur, targets[key], backup_schema, base, suffix)
        stg_rows = _scalar(cur, sql.SQL("select count(*) from {}").format(_ident(stg)))
        expected_latest = _scalar(cur, sql.SQL("select count(distinct inst_key) from {}").format(_ident(stg)))

    conn.autocommit = False
    try:
        with conn.cursor() as cur:
            cur.execute("SET LOCAL lock_timeout = '60s'")
            if stg_rows < 40000:
                raise RuntimeError(f"staging incomplete ({stg_rows} rows < 40000)")
            cur.execute(sql.SQL("TRUNCATE {}").format(_ident(targets["snapshot"])))
            cur.execute(sql.SQL("INSERT INTO {} ({}) SELECT {} FROM {}").format(
                _ident(targets["snapshot"]), col_sql, col_sql, _ident(stg)))
            cur.execute(sql.SQL("TRUNCATE {}").format(_ident(targets["latest"])))
            cur.execute(sql.SQL("INSERT INTO {} ({}) SELECT DISTINCT ON (inst_key) {} FROM {} ORDER BY inst_key, period DESC").format(
                _ident(targets["latest"]), col_sql, col_sql, _ident(targets["snapshot"])))
            cur.execute(sql.SQL("""
                UPDATE {N} t SET target_roa = f.roa, target_efficiency_ratio = f.efficiency_ratio,
                  target_noncurrent_pct = f.noncurrent_assets_pct, target_cof_pct = f.cost_of_funds_pct,
                  target_dep_yoy_financial = f.dep_yoy_pct
                FROM {L} f WHERE f.inst_key = t.target_inst_key
                  AND (t.target_roa, t.target_efficiency_ratio, t.target_noncurrent_pct, t.target_cof_pct, t.target_dep_yoy_financial)
                      IS DISTINCT FROM (f.roa, f.efficiency_ratio, f.noncurrent_assets_pct, f.cost_of_funds_pct, f.dep_yoy_pct)
                """).format(N=_ident(targets["ntt"]), L=_ident(targets["latest"])))
            ntt_updated = cur.rowcount
            snap_n = _scalar(cur, sql.SQL("select count(*) from {}").format(_ident(targets["snapshot"])))
            latest_n = _scalar(cur, sql.SQL("select count(*) from {}").format(_ident(targets["latest"])))
            if snap_n != stg_rows:
                raise RuntimeError(f"post-insert check: snapshot has {snap_n} rows, staging had {stg_rows}")
            if latest_n != expected_latest:
                raise RuntimeError(f"post-insert check: latest has {latest_n} rows, expected {expected_latest}")
        conn.commit()
    except BaseException:
        conn.rollback()
        raise
    finally:
        conn.autocommit = True
    _log(f"  SWAPPED: snapshot {snap_n} rows, latest {latest_n} rows, network_top_targets rows updated {ntt_updated}")
    return {"snapshot": snap_n, "latest": latest_n, "ntt_updated": ntt_updated}


def snapshot_swap():
    _check_guard()
    suffix = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    _log(f"Snapshot refresh SWAP (backups suffixed {suffix}).")
    with _session() as conn:
        with conn.cursor() as cur:
            try:
                new_period, rc_banks = _snapshot_precheck(cur)
                _stage_snapshot(cur)
                _validate_snapshot(cur, new_period, rc_banks).raise_if_failed()
            except BaseException:
                _drop(cur, STG["snapshot"])
                raise
        try:
            _swap_snapshot(conn, STG["snapshot"], PROD, "backup", suffix)
        finally:
            with conn.cursor() as cur:
                _drop(cur, STG["snapshot"])
    _log("Done. Rollback SQL is in the quarterly_refresh module docstring; backups are in schema backup.")


# --------------------------------------------------------------------------
# Step A: UBPR analytics layer
# --------------------------------------------------------------------------

def _stage_ubpr(cur):
    for t in ("peer_stats", "rank", "bank_pg", "coverage"):
        _drop(cur, STG[t])
    _log("  staging analytics.ubpr_peer_stats_clean ...")
    cur.execute(sql.SQL("""
        CREATE TABLE {} AS
        SELECT DISTINCT ON (peer_group, reporting_period, field_code)
               peer_group, peer_group_description, reporting_period, field_code, field_value
        FROM raw.raw_ubpr_peer_stats
        ORDER BY peer_group, reporting_period, field_code, source_file""").format(_ident(STG["peer_stats"])))
    _log("  staging analytics.ubpr_rank_clean (largest; peer_group stays in the key, tier groups only) ...")
    cur.execute(sql.SQL("""
        CREATE TABLE {} AS
        SELECT DISTINCT ON (id_rssd, reporting_period, field_code, peer_group)
               id_rssd, peer_group, reporting_period, field_code, field_value
        FROM raw.raw_ubpr_rank
        WHERE peer_group = ANY(%s)
        ORDER BY id_rssd, reporting_period, field_code, peer_group, source_file""").format(_ident(STG["rank"])),
        (list(TIER_GROUPS),))
    cur.execute(sql.SQL("CREATE TABLE {} AS SELECT DISTINCT id_rssd, reporting_period, peer_group FROM {}")
                .format(_ident(STG["bank_pg"]), _ident(STG["rank"])))
    cur.execute(sql.SQL("""
        CREATE TABLE {cov} AS
        SELECT r.peer_group, r.reporting_period, r.field_code,
               count(distinct r.id_rssd)::bigint AS n_ranked,
               (SELECT count(distinct b.id_rssd) FROM {bpg} b
                 WHERE b.peer_group = r.peer_group AND b.reporting_period = r.reporting_period)::bigint AS n_banks
        FROM {rk} r GROUP BY r.peer_group, r.reporting_period, r.field_code""").format(
        cov=_ident(STG["coverage"]), bpg=_ident(STG["bank_pg"]), rk=_ident(STG["rank"])))
    for t in ("peer_stats", "rank", "bank_pg", "coverage"):
        _lock_down(cur, STG[t])
        _log(f"  staged {STG[t][1]}: {_scalar(cur, sql.SQL('select count(*) from {}').format(_ident(STG[t])))} rows")


def _validate_ubpr(cur, targets=PROD, stg=STG):
    c = Checks()
    for key in ("peer_stats", "rank", "bank_pg", "coverage"):
        s_cols, l_cols = _cols(cur, stg[key]), _cols(cur, targets[key])
        c.add(f"{key}: staging columns == live columns", set(s_cols) == set(l_cols), f"{sorted(set(s_cols) ^ set(l_cols)) or 'match'}")
    for key, label in (("peer_stats", "peer stats"), ("rank", "rank")):
        for period, n in _q(cur, sql.SQL("select reporting_period, count(distinct field_code) from {} group by 1 order by 1").format(_ident(stg[key]))):
            c.add(f"{label}: all {N_CODES} field codes at {period}", n == N_CODES, f"{n} codes")
    for period, n_bpg, n_raw in _q(cur, sql.SQL("""
            select b.reporting_period, b.n, (select count(*) from raw."raw_UBPR" u where u.period = b.reporting_period)
            from (select reporting_period, count(distinct id_rssd) n from {} group by 1) b order by 1""").format(_ident(stg["bank_pg"]))):
        d = abs(n_bpg - n_raw) / n_raw if n_raw else 1
        c.add(f"bank count in peer-group map within 5% of raw_UBPR at {period}", d <= 0.05, f"{n_bpg} vs {n_raw} ({d * 100:.1f}%)")
    multi = _scalar(cur, sql.SQL("select count(*) from (select id_rssd, reporting_period from {} group by 1,2 having count(*) > 1) x").format(_ident(stg["bank_pg"])))
    c.add("each bank has exactly one tier peer group per period", multi == 0, f"{multi} banks with more than one")
    same_periods = _scalar(cur, sql.SQL("""select (select array_agg(distinct reporting_period order by reporting_period) from {s}) =
                                                    (select array_agg(distinct reporting_period order by reporting_period) from {l})""")
                           .format(s=_ident(stg["peer_stats"]), l=_ident(targets["peer_stats"])))
    for key in ("peer_stats", "rank", "bank_pg", "coverage"):
        sn = _scalar(cur, sql.SQL("select count(*) from {}").format(_ident(stg[key])))
        ln = _scalar(cur, sql.SQL("select count(*) from {}").format(_ident(targets[key])))
        if same_periods:
            c.add(f"{key}: no new period, so staging row count reproduces live", sn == ln, f"{sn} vs {ln}")
        else:
            c.info(f"{key}: staging vs live row count (new period present)", f"{sn} vs {ln}")
    return c


def ubpr_layer_dry_run():
    _log("UBPR layer DRY RUN: stage + validate only, no production table is touched.")
    with _session() as conn, conn.cursor() as cur:
        try:
            _stage_ubpr(cur)
            _validate_ubpr(cur).raise_if_failed()
            _log("DRY RUN PASSED: a real swap would validate cleanly on today's data.")
        finally:
            for t in ("peer_stats", "rank", "bank_pg", "coverage"):
                _drop(cur, STG[t])


def _swap_ubpr(conn, stg, targets, backup_schema, suffix):
    keys = ("peer_stats", "rank", "bank_pg", "coverage")
    with conn.cursor() as cur:
        cols = {k: sql.SQL(", ").join(sql.Identifier(c) for c in _cols(cur, targets[k])) for k in keys}
        counts = {k: _scalar(cur, sql.SQL("select count(*) from {}").format(_ident(stg[k]))) for k in keys}
        for k in keys:
            _backup(cur, targets[k], backup_schema, targets[k][1], suffix)
    conn.autocommit = False
    try:
        with conn.cursor() as cur:
            cur.execute("SET LOCAL lock_timeout = '60s'")
            for k in keys:
                if counts[k] == 0:
                    raise RuntimeError(f"staging table for {k} is empty; refusing to truncate")
                cur.execute(sql.SQL("TRUNCATE {}").format(_ident(targets[k])))
                cur.execute(sql.SQL("INSERT INTO {} ({}) SELECT {} FROM {}").format(_ident(targets[k]), cols[k], cols[k], _ident(stg[k])))
                n = _scalar(cur, sql.SQL("select count(*) from {}").format(_ident(targets[k])))
                if n != counts[k]:
                    raise RuntimeError(f"post-insert check: {k} has {n} rows, staging had {counts[k]}")
        conn.commit()
    except BaseException:
        conn.rollback()
        raise
    finally:
        conn.autocommit = True
    _log(f"  SWAPPED: {counts}")
    return counts


def ubpr_layer_swap():
    _check_guard()
    suffix = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    _log(f"UBPR layer SWAP (backups suffixed {suffix}).")
    with _session() as conn:
        try:
            with conn.cursor() as cur:
                _stage_ubpr(cur)
                _validate_ubpr(cur).raise_if_failed()
            _swap_ubpr(conn, STG, PROD, "backup", suffix)
        finally:
            with conn.cursor() as cur:
                for t in ("peer_stats", "rank", "bank_pg", "coverage"):
                    _drop(cur, STG[t])
    _log("Done. Backups are in schema backup; rollback = TRUNCATE + INSERT from them.")


def main():
    args = sys.argv[1:]
    if not args or args[0] not in ("snapshot", "ubpr"):
        sys.exit(__doc__)
    dry = "--dry-run" in args
    {("snapshot", True): snapshot_dry_run, ("snapshot", False): snapshot_swap,
     ("ubpr", True): ubpr_layer_dry_run, ("ubpr", False): ubpr_layer_swap}[(args[0], dry)]()


if __name__ == "__main__":
    main()
