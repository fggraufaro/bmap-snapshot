"""Scheduling entrypoint for a16 -- runs the full pipeline in dependency
order: ingest sources, then rebuild branches_master_v2, then the tiered and
10mi competitor systems, then branch_opportunity_base, then archive the
current year to history.

Steps are defined once in ingestion/pipeline_steps.py -- this script and the
command-center API both drive off that same registry, so there's exactly one
place that defines what the pipeline is and what order it runs in.

Stops at the first failure rather than continuing past it: every rebuild
step reads from what the previous one just wrote, so running e.g. the tiered
system rebuild after branches_master_v2 failed would silently rebuild
against stale data.

GDELT news monitoring is a separate, standalone step (not run here -- see
pipeline_steps.py's docstring) since it can take hours at GDELT's own rate
limit and isn't part of the core ingest-then-rebuild chain.

Usage:
  python -m ingestion.run_all              # run the full pipeline in order
  python -m ingestion.run_all --only STEP_ID   # run a single step
"""

import argparse
import sys
import traceback

from ingestion.pipeline_steps import RUN_ALL_ORDER, STEP_BY_ID


def _run_step(step_id):
    step = STEP_BY_ID[step_id]
    print(f"\n{'=' * 60}\n{step['label']} ({step_id})\n{'=' * 60}")
    step["fn"]()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", type=str, default=None,
                     help=f"run a single step by id (one of {', '.join(RUN_ALL_ORDER)})")
    args, _ = ap.parse_known_args()

    if args.only:
        if args.only not in STEP_BY_ID:
            print(f"Unknown step '{args.only}'. Valid: {', '.join(RUN_ALL_ORDER)}")
            return 1
        _run_step(args.only)
        print(f"\n{args.only}: ok")
        return 0

    results = {}
    for step_id in RUN_ALL_ORDER:
        try:
            _run_step(step_id)
            results[step_id] = "ok"
        except Exception as e:
            print(f"FAILED: {step_id}: {e}")
            traceback.print_exc()
            results[step_id] = f"failed: {e}"
            break  # later steps depend on this one -- don't cascade into stale data

    print(f"\n{'=' * 60}\nSummary\n{'=' * 60}")
    for step_id in RUN_ALL_ORDER:
        print(f"  {step_id}: {results.get(step_id, 'skipped (stopped after earlier failure)')}")

    return 0 if all(v == "ok" for v in results.values()) else 1


if __name__ == "__main__":
    sys.exit(main())
