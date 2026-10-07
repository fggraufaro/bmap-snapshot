"""Rebuild public.network_top_targets (the Hub's "top targets per institution" table).

The database function public.populate_network_top_targets() truncates the table and reloads
it from analytics.branch_target_competitors, ref.dim_institutions and
analytics.bank_financial_snapshot_latest. So run it AFTER the target-competitor rebuild, and after
a ref.dim_institutions refresh (institutions missing from dim_institutions come out with a blank
name).

This wrapper adds what the bare function lacks:
  - a backup of the current table (backup.network_top_targets_<ts>, RLS on),
  - the whole reload inside one transaction, validated before COMMIT (rolled back if the new table
    is empty or shrinks more than 20% versus the old one),
  - the shared advisory lock, so it can't overlap a quarterly swap that also writes this table.
"""

import time

from psycopg2 import sql

from ingestion.quarterly_refresh import _backup, _log, _scalar, _session

TABLE = ("public", "network_top_targets")
MAX_SHRINK = 0.20


def refresh_network_top_targets():
    with _session("refresh_network_top_targets") as conn:
        cur = conn.cursor()
        old_n = _scalar(cur, "select count(*) from public.network_top_targets")
        _backup(cur, TABLE, "backup", "network_top_targets", time.strftime("%Y%m%d_%H%M%S"))
        conn.autocommit = False
        try:
            cur.execute("SELECT public.populate_network_top_targets()")
            new_n = _scalar(cur, "select count(*) from public.network_top_targets")
            blank = _scalar(cur, "select count(*) from public.network_top_targets "
                                 "where my_institution is null or my_institution = ''")
            _log(f"  network_top_targets: {old_n} -> {new_n} rows; {blank} rows with a blank institution name")
            if new_n == 0:
                raise RuntimeError("reload produced 0 rows; rolled back, table unchanged")
            if old_n and new_n < old_n * (1 - MAX_SHRINK):
                raise RuntimeError(f"reload shrank the table from {old_n} to {new_n} rows (more than "
                                   f"{int(MAX_SHRINK * 100)}%); rolled back, table unchanged")
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.autocommit = True
        _log("  committed.")
