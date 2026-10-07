# Refresh plan: every table that feeds a Verlocity tool

Built 2026-10-06 from the tools' own code (bmap-tools HTML, proxy allow-lists, backend scripts) and
the live database (views and RPC functions traced to base tables with pg_depend / function source).
Read-only analysis. "As of" values are what the tables held on that date.

## Tool -> what it reads

| Tool | Direct reads | RPCs |
|---|---|---|
| Intelligence Hub | 10 views/tables | branches_within_radius, radius_market_summary, radius_opportunity_extremes, radius_zip_detail |
| Growth Map | 7 | none called |
| Opportunity View | 8 | branches_within_radius |
| Rate Radar page | 2 | none |
| Assessment doc / Board brief / Snapshot deck (Python) | financial snapshot, opportunity base, target competitors, network_top_targets, dim_institutions, bank_website, persona, Resonate | none |

Views resolved to base tables: vw_branch_opportunity_cbsa, vw_branch_opportunity_yoy_delta,
vw_cfpb_complaints_wow, vw_network_top_targets, vw_prospecting_score, vw_rate_radar_history/latest, vw_zip_persona.

## The 23 base tables and their refresh plan

Legend: **Button** = step in the command center. **Locked** = button exists, destructive part needs ALLOW_ANALYTICS_SWAP.
**GAP** = no refresh path in the pipeline today.

### A. Covered by the main chain ("Run all", steps 1-11)

| Table | As of | Refresh | Cadence |
|---|---|---|---|
| raw.raw_sod | 2026 (306,519 rows) | 1 Ingest FDIC SOD | Annual (SOD publishes ~Jun-Jul) |
| raw.raw_income, raw.raw_population | ACS 2024 | 3 Ingest Census ACS | Annual (new ACS 5-yr ~Dec) |
| raw.raw_zhvi | 2026-07-31 | 4 Ingest Zillow ZHVI | Monthly |
| geo.branches_master_v2 | SOD 2026, CU as of 6/30/2026 | 5 Rebuild branches master (uses 1 + 2) | After any source above |
| geo.branch_competitors_10mi_v2 | 2026-09-25 | 8 Rebuild 10mi | After 5 |
| analytics.branch_opportunity_base | 2026 (94,072) | 9 Rebuild opportunity base | After 5-8 |
| analytics.branch_target_competitors | 265,021 rows (rebuilt 2026-10) | 10 Rebuild target-competitor engine | After 9 |
| analytics.branch_opportunity_base_history | 2025, 2026 | 11 Archive year | Annual |

### B. Button exists but locked or dry-run only

| Table | As of | Step | Why locked |
|---|---|---|---|
| ref.dim_institutions | 8,846 rows, loaded once ~2026-04-02 | 6 Refresh institution directory | Dry run shows 75 adds + 128 renames; apply needs sign-off |
| analytics.bank_financial_snapshot_latest | 2026-06-30 | New quarter: snapshot dry run / refresh | Swap needs ALLOW_ANALYTICS_SWAP and Coordinator sign-off |

### C. GAPS: feed a tool and have no button

| Table | As of | Problem | Proposed fix |
|---|---|---|---|
| public.network_top_targets | 48,241 rows, **247 with blank institution name** | Only a function (`populate_network_top_targets()`, TRUNCATE + reload). No step | Add step after 10; run after dim_institutions apply to fill blanks |
| raw.raw_occupation | ACS 2024 | No loader in the repo at all | Extend census_acs_ingest to the occupation table (S2401), same upsert pattern |
| raw.raw_cfpb_complaints, _trend | 2026-09-11 / 2026-09-18 | Loaded by Session 2; no step here | Wrap Session 2's loader as a Standalone step (needs owner confirmation) |
| geo.uszips, geo."CBSA_zip", ref.bank_website | Static | One-time loads, no step | Annual/semiannual reference refresh (CBSA delineation, zip list); bank_website grows with new institutions |
| public.rate_observations, raw.raw_rate_radar | 2026-10-02 | Rate Radar service refreshes them | Own service on Railway; show last run on the command center, do not duplicate |
| public.persona_runs | 2026-10-05 | Written by the persona layer (Session 3) | Owner: Session 3; event-driven |
| public.resonate_audiences / _enrichment | 2026-10-02 | Fetched by Resonate integration | Owner: Session 2/3; event-driven |

### D. Not a tool-feeding table but part of the new-quarter chain
raw_schedule_RI/RC/UBPR (button), UBPR Stats/Rank (buttons), UBPR analytics layer (locked swap).
raw_schedule_RCE/RIE (read by the Board brief) are stale since 5/29 and not in the loader.

## Command-center grouping (decided 2026-10-06)

- **Group "Ingest all sources"**: SOD, NCUA, Census, ZHVI (each also has its own button).
- **Group "Rebuild BMAP tables"**: branches_master_v2 -> tiered -> 10mi -> opportunity base -> target competitors.
  One button, runs in order, stops at the first failure.
- **Separate buttons** (by hand): institution directory (`ref.dim_institutions`), top targets
  (`public.network_top_targets`, built: backs up, reloads in one transaction, rolls back if empty or >20% smaller),
  yearly archive.
- The old "Run all" button and `/run-all` endpoint are gone.
- **Sequencing catch:** the tiered rebuild loops over `ref.dim_institutions`, so an institution missing from the
  directory gets no tiered rows. The directory therefore has to be refreshed after the branch master and before
  tiered. With the directory as a separate button that is a manual step (run master alone, directory, then the rest).
  Putting the directory in the rebuild group (a dry run until enabled) would remove that manual step.

## Required order after each source change

SOD / NCUA / Census / ZHVI -> branches master -> institution directory -> tiered -> 10mi ->
opportunity base -> target competitors -> top targets -> yearly archive.
New bank quarter: raw load -> UBPR Stats/Rank -> UBPR layer -> financial snapshot; then 9 and 10 only after sign-off.

## Build order proposed
1. network_top_targets step (small, closes a visible Hub gap).
2. Apply dim_institutions (sign-off), then rerun network_top_targets to remove the 247 blanks.
3. Census occupation loader.
4. Wrap CFPB, add RCE/RIE to the quarter loader.
5. Static reference refresh steps (uszips, CBSA_zip, bank_website) and last-run display for Rate Radar / persona / Resonate.

## Caveats
- Direct reads were traced from code; Power BI or other external readers would not appear.
- Growth Map defines an RPC wrapper but no call sites were found.
- "Cadence" is the source's publishing rhythm, not a schedule; nothing runs automatically (buttons only, per Francisco).
