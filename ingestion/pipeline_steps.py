"""Central registry of pipeline steps -- the single source of truth for both
run_all.py (CLI/cron use) and the command-center API (click-to-run UI), so
there's exactly one place that defines what the pipeline is and what order
it runs in.

Two kinds of step:
  - "ingest": calls a Python ingestion script's main()
  - "rpc":    calls a Supabase Postgres function via PostgREST RPC

Rebuild steps must run in this order -- each reads from what the previous
one just wrote:
  ingest sources -> branches_master_v2 -> tiered system -> 10mi system
  -> branch_opportunity_base -> archive to history

GDELT news monitoring is intentionally excluded from RUN_ALL_ORDER -- its
own module docstring notes it can take hours at GDELT's rate limit, and it's
not part of the core ingest-then-rebuild chain. It's a standalone step,
triggered on its own.
"""

import sys
from contextlib import contextmanager

from ingestion import census_acs_ingest, fdic_sod_ingest, gdelt_news_ingest, ncua_fs220_ingest, zhvi_ingest
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
    {"id": "ingest_fdic_sod", "label": "Ingest FDIC SOD (bank branches)",
     "kind": "ingest", "fn": _run_ingest(fdic_sod_ingest.main)},
    {"id": "ingest_ncua", "label": "Ingest NCUA (CU branches + FS220)",
     "kind": "ingest", "fn": _run_ingest(ncua_fs220_ingest.main)},
    {"id": "ingest_census", "label": "Ingest Census ACS (income/population)",
     "kind": "ingest", "fn": _run_ingest(census_acs_ingest.main)},
    {"id": "ingest_zhvi", "label": "Ingest Zillow ZHVI (home values)",
     "kind": "ingest", "fn": _run_ingest(zhvi_ingest.main)},
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
    {"id": "archive_year", "label": "Archive current year to history",
     "kind": "rpc", "fn": _archive_year},
]

# Standalone -- not part of RUN_ALL_ORDER (see module docstring).
GDELT_STEP = {"id": "ingest_gdelt", "label": "Ingest GDELT competitor news",
              "kind": "ingest", "fn": _run_ingest(gdelt_news_ingest.main)}

ALL_STEPS = STEPS + [GDELT_STEP]
STEP_BY_ID = {s["id"]: s for s in ALL_STEPS}
RUN_ALL_ORDER = [s["id"] for s in STEPS]
