"""Scheduling entrypoint for a16 — runs every bulk ingestion script, then
refreshes the scored views once.

Design: each ingestion script (zhvi_ingest, ncua_fs220_ingest,
fdic_sod_ingest, census_acs_ingest) is already idempotent and
self-limiting (auto-detects the latest available period/vintage and
no-ops or re-upserts unchanged data if nothing's new), so it's safe to
run this daily via a Railway cron-scheduled service rather than wiring
four separate schedules for four different real-world cadences
(monthly/quarterly/quarterly/annual). One run, one log, one place to
look when something breaks.

The refresh (analytics.refresh_branch_opportunity_base — confirmed via
information_schema.routines as the current one: the tiered-adaptive-
radius, 16-play-matrix version referenced in this repo's recent commits,
not the older public.refresh_branch_opportunity_base) runs ONCE at the
end regardless of which sources actually had new data, not after each
script individually — it's a full ~93k-branch rebuild, so doing it once
per run instead of up to 4x saves real time and load.

NOT wired into an actual Railway cron schedule yet — see a16 in
SESSION_COORDINATION.md. This script is what that schedule should run.

Usage:
  python -m ingestion.run_all              # run everything, then refresh
  python -m ingestion.run_all --no-refresh # skip the refresh step (for testing ingestion alone)
"""

import argparse
import sys
import traceback

from contextlib import contextmanager

from ingestion import census_acs_ingest, fdic_sod_ingest, ncua_fs220_ingest, zhvi_ingest
from ingestion.supabase_client import call_rpc

JOBS = [
    ("zhvi", zhvi_ingest.main),
    ("ncua_fs220", ncua_fs220_ingest.main),
    ("fdic_sod", fdic_sod_ingest.main),
    ("census_acs", census_acs_ingest.main),
]


@contextmanager
def _bare_argv():
    """Each job's own main() runs argparse.parse_args() against sys.argv,
    which would otherwise see run_all's own flags (e.g. --no-refresh) and
    reject them as unrecognized. Jobs are always run with their defaults
    here, so just hide argv for the duration of the call."""
    saved = sys.argv
    sys.argv = [saved[0]]
    try:
        yield
    finally:
        sys.argv = saved


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-refresh", action="store_true", help="skip analytics.refresh_branch_opportunity_base at the end")
    args, _ = ap.parse_known_args()

    results = {}
    for name, job in JOBS:
        print(f"\n{'=' * 60}\n{name}\n{'=' * 60}")
        try:
            with _bare_argv():
                job()
            results[name] = "ok"
        except Exception as e:
            print(f"FAILED: {name}: {e}")
            traceback.print_exc()
            results[name] = f"failed: {e}"

    print(f"\n{'=' * 60}\nSummary\n{'=' * 60}")
    for name, status in results.items():
        print(f"  {name}: {status}")

    if args.no_refresh:
        print("\n--no-refresh set, skipping analytics.refresh_branch_opportunity_base.")
        return

    print("\nRefreshing analytics.refresh_branch_opportunity_base()...")
    call_rpc("refresh_branch_opportunity_base", schema="analytics")
    print("Refresh complete.")


if __name__ == "__main__":
    sys.exit(main())
