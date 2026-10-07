"""Central registry of pipeline steps -- the single source of truth for both
run_all.py (CLI/cron use) and the command-center API (click-to-run UI), so
there's exactly one place that defines what the pipeline is and what order
it runs in.

Two kinds of step:
  - "ingest": calls a Python ingestion script's main()
  - "rpc":    calls a Supabase Postgres function via PostgREST RPC

Steps are grouped (see GROUPS at the bottom): "sources" and "rebuild" each run
as one sequenced button. The rebuild group must run in this order -- each step
reads from what the previous one just wrote:
  branches_master_v2 -> tiered system -> 10mi system -> branch_opportunity_base
  -> branch_target_competitors
ref.dim_institutions, public.network_top_targets and the yearly archive are
separate buttons, run by hand after the rebuild group.

GDELT news monitoring is intentionally excluded from every group -- its
own module docstring notes it can take hours at GDELT's rate limit, and it's
not part of the core ingest-then-rebuild chain. It's a standalone step,
triggered on its own.
"""

import sys
from contextlib import contextmanager

from ingestion import quarterly_refresh, refresh_dim_institutions, refresh_network_top_targets, session2_datasets
from ingestion import census_acs_ingest, fdic_sod_ingest, ffiec_call_report_ingest, ffiec_ubpr_peer_stats_ingest, ffiec_ubpr_rank_ingest, gdelt_news_ingest, ncua_fs220_ingest, zhvi_ingest
from ingestion.pg_direct import call_procedure
from ingestion.supabase_client import get


@contextmanager
def _bare_argv():
    """Run a job's own main() with its own defaults, hiding whatever argv
    the calling process (run_all.py, or the API server) was started with --
    each script's own argparse would otherwise choke on unrecognized flags."""
    saved = sys.argv
    sys.argv = [saved[0]]
    try:
        yield
    finally:
        sys.argv = saved


def _run_ingest(fn):
    def _call():
        with _bare_argv():
            fn()
    return _call


def latest_sod_year():
    """The most recently ingested FDIC SOD year -- used as the target year
    for branches_master_v2 / archive, so nobody has to hardcode or manually
    bump a year value in code every year. raw.raw_sod's YEAR column is text,
    but always a 4-digit year, so lexicographic ordering matches numeric
    ordering."""
    rows = get("raw_sod", schema="raw", params='select=%22YEAR%22&order=%22YEAR%22.desc&limit=1')
    if not rows:
        raise RuntimeError("raw.raw_sod is empty -- ingest FDIC SOD data first")
    return int(rows[0]["YEAR"])


def _rebuild_branches_master():
    year = latest_sod_year()
    call_procedure("CALL geo.refresh_branches_master_v2(%s)", (year,))


def _archive_year():
    year = latest_sod_year()
    call_procedure("SELECT public.archive_bmap_year_snapshot(%s)", (year,))


STEPS = [
    # Group "sources": each source publishes on its own schedule, so each has its own button;
    # the group button runs them one after another.
    {"id": "ingest_fdic_sod", "label": "Ingest FDIC SOD (bank branches)",
     "kind": "ingest", "fn": _run_ingest(fdic_sod_ingest.main)},
    {"id": "ingest_ncua", "label": "Ingest NCUA (CU branches + FS220)",
     "kind": "ingest", "fn": _run_ingest(ncua_fs220_ingest.main)},
    {"id": "ingest_census", "label": "Ingest Census ACS (income/population)",
     "kind": "ingest", "fn": _run_ingest(census_acs_ingest.main)},
    {"id": "ingest_zhvi", "label": "Ingest Zillow ZHVI (home values)",
     "kind": "ingest", "fn": _run_ingest(zhvi_ingest.main)},
    # Group "rebuild": each step reads what the one before it just wrote.
    {"id": "rebuild_branches_master", "label": "Rebuild branches_master_v2",
     "kind": "rpc", "fn": _rebuild_branches_master},
    {"id": "rebuild_tiered", "label": "Rebuild tiered competitor system",
     "kind": "rpc", "fn": lambda: call_procedure("CALL public.refresh_tiered_competitor_system()")},
    {"id": "rebuild_10mi", "label": "Rebuild 10mi competitor system (capped top-50)",
     "kind": "rpc", "fn": lambda: call_procedure("CALL public.refresh_old_10mi_competitor_system()")},
    {"id": "rebuild_opportunity_base", "label": "Rebuild branch_opportunity_base",
     "kind": "rpc", "fn": lambda: call_procedure("CALL public.refresh_branch_opportunity_base_with_backup()")},
    {"id": "rebuild_target_competitors", "label": "Rebuild target-competitor engine (Hub drill-downs)",
     "kind": "rpc", "fn": lambda: call_procedure("CALL analytics.refresh_branch_target_competitors()")},
]

# Separate buttons, deliberately NOT in a group. Run them by hand, in the order the page shows.
# Reads the branch master (to keep AK/HI-headquartered institutions that have scored branches).
# Applies only when ALLOW_ANALYTICS_SWAP=yes; otherwise it does a dry run. See refresh_dim_institutions.py.
DIM_STEP = {"id": "refresh_dim_institutions", "label": "Refresh institution directory (ref.dim_institutions)",
            "kind": "refresh", "fn": _run_ingest(refresh_dim_institutions.refresh_dim_institutions),
            "guard_fallback": quarterly_refresh.GUARD_ENV}
# Reads branch_target_competitors + ref.dim_institutions + the financial snapshot; truncates and
# reloads in one transaction, with a backup first. See refresh_network_top_targets.py.
NTT_STEP = {"id": "refresh_network_top_targets", "label": "Refresh top targets per institution (public.network_top_targets)",
            "kind": "refresh", "fn": _run_ingest(refresh_network_top_targets.refresh_network_top_targets)}
ARCHIVE_STEP = {"id": "archive_year", "label": "Archive current year to history",
                "kind": "rpc", "fn": _archive_year}

# Group "datasets": the reference datasets Session 2 loaded by hand. They feed the opportunity-base rebuild
# (CBP, Business Formation, IRS migration), the persona brief (QCEW) and Growth Map (CFPB trend), so run them
# BEFORE the rebuild group. Each one reports a diff against the live table and applies only when
# ALLOW_DATASET_LOAD=yes (separate from ALLOW_ANALYTICS_SWAP). See session2_datasets.py.
_DS = session2_datasets
DATASET_STEPS = [
    {"id": "refresh_cbp", "label": "Refresh Census CBP ZIP business counts (SMB index)",
     "kind": "refresh", "fn": _run_ingest(_DS.refresh_cbp), "guard_fallback": _DS.GUARD_ENV},
    {"id": "refresh_business_formation", "label": "Refresh Census Business Formation (state)",
     "kind": "refresh", "fn": _run_ingest(_DS.refresh_business_formation), "guard_fallback": _DS.GUARD_ENV},
    {"id": "refresh_irs_migration", "label": "Refresh IRS state migration",
     "kind": "refresh", "fn": _run_ingest(_DS.refresh_irs_migration), "guard_fallback": _DS.GUARD_ENV},
    {"id": "refresh_qcew", "label": "Refresh BLS QCEW state wages (persona brief)",
     "kind": "refresh", "fn": _run_ingest(_DS.refresh_qcew), "guard_fallback": _DS.GUARD_ENV},
    {"id": "refresh_cfpb_trend", "label": "Refresh CFPB complaint trend (Growth Map, weekly)",
     "kind": "refresh", "fn": _run_ingest(_DS.refresh_cfpb_trend), "guard_fallback": _DS.GUARD_ENV},
]

# Standalone -- not part of any group (see module docstring).
GDELT_STEP = {"id": "ingest_gdelt", "label": "Ingest GDELT competitor news",
              "kind": "ingest", "fn": _run_ingest(gdelt_news_ingest.main)}

# Standalone -- feeds a84 (UBPR Peer-Benchmarking), not the branches_master_v2
# rebuild chain. Session 3 owns the mapping/wiring once this lands.
UBPR_PEER_STATS_STEP = {"id": "ingest_ubpr_peer_stats", "label": "Ingest FFIEC UBPR peer-group stats (a84)",
                         "kind": "ingest", "fn": _run_ingest(ffiec_ubpr_peer_stats_ingest.main)}

# Standalone, same reason as above. Large (~30M rows, institution x peer
# group x field) -- expect this one to run considerably longer than the
# Stats pull.
UBPR_RANK_STEP = {"id": "ingest_ubpr_rank", "label": "Ingest FFIEC UBPR bank-vs-peer rank (a84)",
                   "kind": "ingest", "fn": _run_ingest(ffiec_ubpr_rank_ingest.main)}

# Standalone, and deliberately NOT in any group (a85). Loads the newest
# FFIEC quarter into raw_schedule_RI/RC + raw_UBPR (raw tables only; skips
# any table that already has it, errors on a partial load; never deletes).
# It does NOT run refresh_bmap_after_upload, a database procedure that
# truncates and rebuilds the analytics snapshot/score tables. Nothing in this
# registry calls it, so it cannot be triggered from the command center. Loading
# a quarter here only adds raw rows; the snapshot refresh is a separate, manual
# step (see the staged swap Session 3 is specifying for the 9/30 load).
CALL_REPORT_STEP = {"id": "ingest_bank_quarter", "label": "Ingest latest FFIEC bank quarter (RI/RC/UBPR, raw only)",
                     "kind": "ingest", "fn": _run_ingest(ffiec_call_report_ingest.load_latest)}

# New-quarter refresh (a84): run in this order when FFIEC publishes a quarter.
# The two *_swap steps are DESTRUCTIVE (back up, then TRUNCATE + INSERT the
# production analytics tables), so they carry "guard_env": they refuse to run,
# and the page shows them locked, until ALLOW_ANALYTICS_SWAP=yes is set on the
# service. It is deliberately not set (a90 hold; Coordinator sign-off first).
# The dry-run steps stage + validate + report and never touch a production
# table. None of these call refresh_bmap_after_upload or rebuild
# branch_opportunity_base / branch_target_competitors.
UBPR_LAYER_DRY_RUN_STEP = {"id": "ubpr_layer_dry_run", "label": "UBPR analytics layer: dry run (stage + validate, changes nothing)",
                            "kind": "refresh", "fn": _run_ingest(quarterly_refresh.ubpr_layer_dry_run)}
UBPR_LAYER_SWAP_STEP = {"id": "ubpr_layer_swap", "label": "UBPR analytics layer: rebuild (backs up, then swaps)",
                         "kind": "refresh", "fn": _run_ingest(quarterly_refresh.ubpr_layer_swap),
                         "guard_env": quarterly_refresh.GUARD_ENV}
SNAPSHOT_DRY_RUN_STEP = {"id": "snapshot_dry_run", "label": "Financial snapshot: dry run (stage + validate, changes nothing)",
                          "kind": "refresh", "fn": _run_ingest(quarterly_refresh.snapshot_dry_run)}
SNAPSHOT_SWAP_STEP = {"id": "snapshot_swap", "label": "Financial snapshot: refresh (backs up, then swaps)",
                       "kind": "refresh", "fn": _run_ingest(quarterly_refresh.snapshot_swap),
                       "guard_env": quarterly_refresh.GUARD_ENV}

ALL_STEPS = STEPS + DATASET_STEPS + [DIM_STEP, NTT_STEP, ARCHIVE_STEP, GDELT_STEP, UBPR_PEER_STATS_STEP, UBPR_RANK_STEP, CALL_REPORT_STEP,
                     UBPR_LAYER_DRY_RUN_STEP, UBPR_LAYER_SWAP_STEP, SNAPSHOT_DRY_RUN_STEP, SNAPSHOT_SWAP_STEP]
STEP_BY_ID = {s["id"]: s for s in ALL_STEPS}

# Groups: one button runs the group's steps in order, stopping at the first failure.
GROUPS = [
    {"id": "sources", "label": "Ingest all sources",
     "note": "Pulls SOD, NCUA, Census and ZHVI one after another. Each source also has its own button.",
     "steps": ["ingest_fdic_sod", "ingest_ncua", "ingest_census", "ingest_zhvi"]},
    {"id": "datasets", "label": "Refresh other datasets",
     "note": "CBP, Business Formation, IRS migration, QCEW and the CFPB trend. Run before the rebuild group: the first "
             "three feed the opportunity base. Each step shows its changes against the live table and writes nothing "
             "until ALLOW_DATASET_LOAD is enabled on the service.",
     "steps": [s["id"] for s in DATASET_STEPS]},
    {"id": "rebuild", "label": "Rebuild BMAP tables",
     "note": "Branch master, tiered, 10mi, opportunity base, then target competitors, in that order. "
             "Takes a while (the last two alone are several minutes).",
     "steps": ["rebuild_branches_master", "rebuild_tiered", "rebuild_10mi",
               "rebuild_opportunity_base", "rebuild_target_competitors"]},
]
GROUP_BY_ID = {g["id"]: g for g in GROUPS}
# Kept for run_all.py (CLI): sources then rebuild, nothing else.
RUN_ALL_ORDER = [sid for g in GROUPS for sid in g["steps"]]

# Display order for the "new-quarter refresh" section of the command center.
QUARTERLY_ORDER = ["ingest_bank_quarter", "ingest_ubpr_peer_stats", "ingest_ubpr_rank",
                   "ubpr_layer_dry_run", "ubpr_layer_swap", "snapshot_dry_run", "snapshot_swap"]
