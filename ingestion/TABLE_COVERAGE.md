# Table coverage control document

Generated 2026-10-06 by `python -m ingestion.table_coverage --write` from the live database and `ingestion/table_coverage.json`. Do not edit by hand: edit the JSON, re-run, commit both.

**Result: PASS** - 0 errors, 17 warnings. 102 objects in the database (81 tables, 21 views) all have a row below; the backup schema (34 tables) is exempt.

## Summary

| Status | Objects | Meaning |
|---|---|---|
| GAP | 7 | Feeds a tool or rebuild, no refresh path |
| STATIC | 4 | Loaded once, no refresh path |
| UNUSED | 12 | No reader, no refresh: keep or drop |
| LOCKED | 7 | Button exists, destructive part gated |
| HELD | 1 | Button exists, deliberately not clicked |
| AUTO | 22 | Refreshed by a command-center button |
| ARCHIVE | 6 | History, yearly archive button |
| SERVICE | 9 | Refreshed by another service |
| APP | 10 | Written by an application |
| DERIVED | 21 | View, follows its base tables |
| SYSTEM | 3 | PostGIS / platform |

## Errors

None.

## Warnings (open items)

- RLS OFF: analytics.branch_target_competitors_backup_pre_a55 (anon/authenticated may read it unfiltered if granted)
- RLS OFF: analytics.branch_target_competitors_backup_pre_a55_v2 (anon/authenticated may read it unfiltered if granted)
- FEEDS A TOOL, NO PIPELINE REFRESH: geo.CBSA_zip [STATIC] -> Intelligence Hub, Opportunity View
- FEEDS A TOOL, NO PIPELINE REFRESH: geo.uszips [STATIC] -> Growth Map, Intelligence Hub
- FEEDS A TOOL, NO PIPELINE REFRESH: raw.raw_cfpb_complaints [GAP] -> Growth Map
- FEEDS A TOOL, NO PIPELINE REFRESH: raw.raw_cfpb_complaints_trend [GAP] -> Growth Map
- FEEDS A TOOL, NO PIPELINE REFRESH: raw.raw_occupation [GAP] -> Bank Assessment doc, Growth Map, Intelligence Hub, Opportunity View
- FEEDS A TOOL, NO PIPELINE REFRESH: ref.bank_website [STATIC] -> Board Brief, Intelligence Hub, Snapshot deck
- VIEW DEPENDS ON UNREFRESHED TABLE: public.vw_branch_opportunity_cbsa <- geo.CBSA_zip
- VIEW DEPENDS ON UNREFRESHED TABLE: public.vw_zip_persona <- raw.raw_occupation
- VIEW DEPENDS ON UNREFRESHED TABLE: public.vw_cfpb_complaints_wow <- raw.raw_cfpb_complaints, raw.raw_cfpb_complaints_trend
- VIEW DEPENDS ON UNREFRESHED TABLE: public.vw_smb_index_by_zip <- raw.raw_business_formation_state, raw.raw_cbp_totals
- VIEW DEPENDS ON UNREFRESHED TABLE: geo.dim_zip_cbsa <- geo.CBSA_zip
- VIEW DEPENDS ON UNREFRESHED TABLE: public.vw_bank_directory <- ref.bank_website
- REBUILD INPUT NOT REFRESHED: analytics.refresh_branch_opportunity_base reads raw.raw_business_formation_state [GAP]
- REBUILD INPUT NOT REFRESHED: analytics.refresh_branch_opportunity_base reads raw.raw_irs_migration_state [GAP]
- REBUILD INPUT NOT REFRESHED: analytics.refresh_branch_opportunity_base reads ref.banks [STATIC]

## Rebuild chain: inputs of each procedure

| Procedure | Reads (status) |
|---|---|
| `geo.refresh_branches_master_v2` | `geo.branches_master_v2` AUTO, `raw.Raw_cu_fs220` AUTO, `raw.raw_cu_branches` AUTO, `raw.raw_sod` AUTO |
| `analytics.rebuild_tiered_radius_batch` | `analytics.branch_opportunity_base` AUTO, `geo.branch_competitors_tiered_v1` AUTO, `geo.branches_master_v2` AUTO, `ref.dim_institutions` LOCKED |
| `analytics.refresh_branch_competitors_tiered_v1` | `geo.branch_competitors_tiered_v1` AUTO, `geo.branches_master_v2` AUTO |
| `geo.refresh_branch_competitors_10mi_v2` | `geo.branch_competitors_10mi_v2` AUTO, `geo.branches_master_v2` AUTO |
| `analytics.refresh_branch_opportunity_base` | `analytics.bank_financial_snapshot_latest` LOCKED, `analytics.branch_opportunity_base` AUTO, `analytics.branch_opportunity_minmax` AUTO, `geo.branch_competitors_tiered_v1` AUTO, `geo.branch_radius_stats_tiered_v1` AUTO, `geo.branches_master_v2` AUTO, `public.vw_smb_index_by_zip` DERIVED, `raw.raw_business_formation_state` GAP, `raw.raw_income` AUTO, `raw.raw_irs_migration_state` GAP, `raw.raw_population` AUTO, `raw.raw_sod` AUTO, `raw.raw_zhvi` AUTO, `ref.banks` STATIC |
| `analytics.refresh_branch_target_competitors` | `analytics.branch_opportunity_base` AUTO, `analytics.branch_target_competitors` AUTO, `geo.branches_master_v2` AUTO |
| `public.populate_network_top_targets` | `analytics.bank_financial_snapshot_latest` LOCKED, `analytics.branch_opportunity_base` AUTO, `analytics.branch_target_competitors` AUTO, `public.network_top_targets` AUTO, `ref.dim_institutions` LOCKED |

## Every object

### GAP (7)

| Object | Rows | Newest data | How it is refreshed | Owner | Cadence | Feeds | Action / note |
|---|---|---|---|---|---|---|---|
| `raw.raw_business_formation_state` | 51 |  | none (loaded by Session 2, no step) | Session 2 | Annual |  | Wrap the loader as a standalone step - Feeds the opportunity-base rebuild (via vw_smb_index_by_zip), so every tool |
| `raw.raw_cbp_totals` | 34,954 |  | none (one-time load) | Session 2 | Annual |  | Wrap the loader as a standalone step - Feeds the opportunity-base rebuild via vw_smb_index_by_zip |
| `raw.raw_cfpb_complaints` | 53 | 2026-09-11 | none (loaded by Session 2, no step) | Session 2 | Weekly/monthly | Growth Map | Wrap the loader as a standalone step - Feeds Growth Map via vw_cfpb_complaints_wow |
| `raw.raw_cfpb_complaints_trend` | 106 | 2026-09-18 | none (loaded by Session 2, no step) | Session 2 | Weekly/monthly | Growth Map | Wrap the loader as a standalone step - Feeds Growth Map via vw_cfpb_complaints_wow |
| `raw.raw_irs_migration_state` | 51 |  | none (loaded by Session 2, no step) | Session 2 | Annual |  | Wrap the loader as a standalone step - Feeds the opportunity-base rebuild directly |
| `raw.raw_occupation` | 101,318 | 2024 | none (no loader exists) | Session 1 | Annual (ACS ~Dec) | Bank Assessment doc, Growth Map, Intelligence Hub, Opportunity View | Extend the Census loader to the occupation table - Stuck at ACS 2024 |
| `raw.raw_schedule_RCE` | 22,224 |  | none (stale since 5/29) | Session 1 | Quarterly (FFIEC) |  | Add RCE to the call-report loader - Read by the Board brief |

### STATIC (4)

| Object | Rows | Newest data | How it is refreshed | Owner | Cadence | Feeds | Action / note |
|---|---|---|---|---|---|---|---|
| `geo.CBSA_zip` | 47,634 |  | none (one-time load) | Session 1 | Yearly (CBSA delineation changes) | Intelligence Hub, Opportunity View | Add a reference-refresh step, or document the source and a yearly review |
| `geo.uszips` | 33,791 |  | none (one-time load) | Session 1 | Yearly | Growth Map, Intelligence Hub | Add a reference-refresh step, or document the source and a yearly review |
| `ref.bank_website` | 4,403 |  | none (one-time load) | Session 1 | As institutions are added | Board Brief, Intelligence Hub, Snapshot deck | Add a refresh step, or tie it to the institution directory - Feeds Hub, Board brief, Snapshot deck |
| `ref.banks` | 1,715 |  | none found in code | Session 1 | Unverified |  | Owner to confirm how it is loaded - Read by the opportunity-base rebuild and many scripts; I found no loader |

### UNUSED (12)

| Object | Rows | Newest data | How it is refreshed | Owner | Cadence | Feeds | Action / note |
|---|---|---|---|---|---|---|---|
| `analytics.branch_cu_exposure` | 952,902 |  | none | Session 3 | - |  | Decide keep or drop (953K rows) - No loader, no reader in code or in any view/function |
| `analytics.branch_target_competitors_backup_20260919` | 253,788 |  | none (stray backup kept in analytics) | Session 1 | - |  | Move to backup schema or drop - Superseded by the a90 rebuild |
| `analytics.branch_target_competitors_backup_pre_a55` | 264,101 |  | none (stray backup kept in analytics) | Session 1 | - |  | Enable RLS now (it is off), then move to backup schema or drop - Superseded by the a90 rebuild |
| `analytics.branch_target_competitors_backup_pre_a55_v2` | 256,151 |  | none (stray backup kept in analytics) | Session 1 | - |  | Enable RLS now (it is off), then move to backup schema or drop - Superseded by the a90 rebuild |
| `analytics.cu_institution_scores` | 4,624 |  | none | Session 3 | - |  | Decide keep or drop (4,624 rows) - No loader, no reader |
| `analytics.institution_structure_changes` | 159 |  | none | Session 2 | - |  | Session 2 to confirm purpose or drop - No loader, no reader (159 rows) |
| `pbi.branches_master` | 72,998 |  | none (Power BI legacy copy) | Session 1 | - |  | Confirm whether Power BI still reads it - 72,998 rows vs 94,076 in geo.branches_master_v2; stale copy |
| `raw.raw_fdic_structure_changes` | 1,273 |  | none | Session 2 | Quarterly |  | Session 2 to confirm it will be read, or drop - No reader yet (1,273 rows) |
| `raw.raw_qcew_state` | 51 |  | none | Session 2 | Quarterly |  | Session 2 to confirm it will be read, or drop - No reader yet |
| `raw.raw_schedule_RIE` | 22,241 |  | none | Session 1 | Quarterly (FFIEC) |  | Add to the loader with RCE, or drop - Read only by refresh_bmap_after_upload, which is never run |
| `raw.raw_sec_8k_filings` | 30 |  | none | Session 2 | - |  | Drop (empty, no reader) or build the loader - 0 rows |
| `ref.banks_stage` | 1,715 |  | none | Session 1 | - |  | Drop (staging leftover, no reader) |

### LOCKED (7)

| Object | Rows | Newest data | How it is refreshed | Owner | Cadence | Feeds | Action / note |
|---|---|---|---|---|---|---|---|
| `analytics.bank_financial_snapshot` | 62,437 | 2026-06-30 | button: snapshot_swap (new-quarter section) | Session 1 | Quarterly |  | Swap gated by ALLOW_ANALYTICS_SWAP + Coordinator sign-off |
| `analytics.bank_financial_snapshot_latest` | 9,272 | 2026-06-30 | button: snapshot_swap (new-quarter section) | Session 1 | Quarterly | Bank Assessment doc, Board Brief, Intelligence Hub, Opportunity View, Snapshot deck | Read by the opportunity-base rebuild: run that after the snapshot swap |
| `analytics.ubpr_bank_peer_group` | 8,565 |  | button: ubpr_layer_swap (new-quarter section) | Session 1 | Quarterly |  | UBPR analytics layer (bank-to-peer-group map) |
| `analytics.ubpr_peer_stats_clean` | 48,994 |  | button: ubpr_layer_swap (new-quarter section) | Session 1 | Quarterly |  | UBPR analytics layer (peer stats) |
| `analytics.ubpr_rank_clean` | 1,047,749 |  | button: ubpr_layer_swap (new-quarter section) | Session 1 | Quarterly |  | UBPR analytics layer (ranks) |
| `analytics.ubpr_rank_coverage` | 4,349 |  | button: ubpr_layer_swap (new-quarter section) | Session 1 | Quarterly |  | UBPR analytics layer (rank coverage) |
| `ref.dim_institutions` | 8,846 |  | button: refresh_dim_institutions (separate; dry run until enabled) | Session 1 | After branch master, before tiered | Board Brief, Growth Map, Intelligence Hub, Opportunity View, Snapshot deck | Apply needs ALLOW_ANALYTICS_SWAP and sign-off; the tiered rebuild loops over this table |

### HELD (1)

| Object | Rows | Newest data | How it is refreshed | Owner | Cadence | Feeds | Action / note |
|---|---|---|---|---|---|---|---|
| `raw.raw_gdelt_news` | 0 |  | button: ingest_gdelt (standalone, held a87) | Session 1 | On demand |  | Stays unclicked until a87 clears |

### AUTO (22)

| Object | Rows | Newest data | How it is refreshed | Owner | Cadence | Feeds | Action / note |
|---|---|---|---|---|---|---|---|
| `analytics.branch_opportunity_base` | 94,072 | 2026 | button: rebuild_opportunity_base (rebuild group) | Session 1 | After any source change | Bank Assessment doc, Board Brief, Growth Map, Intelligence Hub, Opportunity View, Snapshot deck |  |
| `analytics.branch_opportunity_minmax` | 52 |  | written by the rebuild_opportunity_base procedure | Session 1 | With opportunity base |  |  |
| `analytics.branch_target_competitors` | 265,021 |  | button: rebuild_target_competitors (rebuild group) | Session 1 | After opportunity base | Bank Assessment doc, Board Brief, Intelligence Hub, Opportunity View, Snapshot deck |  |
| `geo.branch_competitors_10mi_v2` | 3,428,648 | 2026-09-25 | button: rebuild_10mi (rebuild group) | Session 1 | After branch master | Growth Map |  |
| `geo.branch_competitors_tiered_v1` | 3,928,985 |  | button: rebuild_tiered (rebuild group) | Session 1 | After branch master and institution directory |  |  |
| `geo.branch_radius_stats_10mi_v2` | 94,076 |  | button: rebuild_10mi (rebuild group) | Session 1 | After branch master |  |  |
| `geo.branch_radius_stats_tiered_v1` | 94,076 |  | button: rebuild_tiered (rebuild group) | Session 1 | After branch master and institution directory |  |  |
| `geo.branches_master_v2` | 94,076 | 2026-06-30 | button: rebuild_branches_master (rebuild group) | Session 1 | After SOD / NCUA load | Bank Assessment doc, Intelligence Hub, Opportunity View |  |
| `public.network_top_targets` | 48,241 |  | button: refresh_network_top_targets (separate) | Session 1 | After target competitors | Bank Assessment doc, Board Brief, Intelligence Hub, Snapshot deck | Also updated in place by snapshot_swap (financial columns) |
| `raw.Raw_cu_fs220` | 17,650 | 2026-06-30 | button: ingest_ncua (sources group) | Session 1 | Quarterly (NCUA) |  |  |
| `raw.raw_UBPR` | 44,789 | 2026-06-30 | button: ingest_bank_quarter (new-quarter section) | Session 1 | Quarterly (FFIEC) |  |  |
| `raw.raw_cu_branches` | 44,905 | 2026-06-30 | button: ingest_ncua (sources group) | Session 1 | Quarterly (NCUA) |  |  |
| `raw.raw_income` | 135,092 | 2024 | button: ingest_census (sources group) | Session 1 | Annual (ACS ~Dec) | Growth Map, Intelligence Hub, Opportunity View |  |
| `raw.raw_population` | 135,092 | 2024 | button: ingest_census (sources group) | Session 1 | Annual (ACS ~Dec) | Bank Assessment doc, Growth Map, Intelligence Hub, Opportunity View |  |
| `raw.raw_schedule_RC` | 44,787 | 2026-06-30 | button: ingest_bank_quarter (new-quarter section) | Session 1 | Quarterly (FFIEC) |  |  |
| `raw.raw_schedule_RI` | 44,787 | 2026-06-30 | button: ingest_bank_quarter (new-quarter section) | Session 1 | Quarterly (FFIEC) |  |  |
| `raw.raw_sod` | 306,519 | 2026 | button: ingest_fdic_sod (sources group) | Session 1 | Yearly (SOD ~Jun-Jul) | Bank Assessment doc, Intelligence Hub |  |
| `raw.raw_ubpr_peer_stats` | 65,450 |  | button: ingest_ubpr_peer_stats (new-quarter section) | Session 1 | Quarterly |  |  |
| `raw.raw_ubpr_rank` | 5,906,065 |  | button: ingest_ubpr_rank (new-quarter section) | Session 1 | Quarterly |  |  |
| `raw.raw_zhvi` | 1,124,974 | 2026-07-31 | button: ingest_zhvi (sources group) | Session 1 | Monthly | Intelligence Hub |  |
| `raw.ref_ubpr_rank_fields` | 131 |  | written by the UBPR Rank ingest step | Session 1 | Quarterly |  |  |
| `raw.ref_ubpr_stats_fields` | 1,260 |  | written by the UBPR Stats ingest step | Session 1 | Quarterly |  |  |

### ARCHIVE (6)

| Object | Rows | Newest data | How it is refreshed | Owner | Cadence | Feeds | Action / note |
|---|---|---|---|---|---|---|---|
| `analytics.branch_opportunity_base_history` | 187,676 | 2025, 2026 | button: archive_year (separate) | Session 1 | Yearly | Growth Map |  |
| `geo.branch_competitors_10mi_v2_history` | 6,668,801 |  | button: archive_year (separate) | Session 1 | Yearly |  |  |
| `geo.branch_competitors_tiered_v1_history` | 7,899,393 |  | button: archive_year (separate) | Session 1 | Yearly |  |  |
| `geo.branch_radius_stats_10mi_v2_history` | 94,076 |  | button: archive_year (separate) | Session 1 | Yearly |  |  |
| `geo.branch_radius_stats_tiered_v1_history` | 187,680 |  | button: archive_year (separate) | Session 1 | Yearly |  |  |
| `geo.branches_master_v2_history` | 187,684 |  | button: archive_year (separate) | Session 1 | Yearly |  |  |

### SERVICE (9)

| Object | Rows | Newest data | How it is refreshed | Owner | Cadence | Feeds | Action / note |
|---|---|---|---|---|---|---|---|
| `public.bank_registry` | 133 |  | Rate Radar service | Rate Radar | Per Rate Radar run |  |  |
| `public.extraction_cache` | 103 |  | Rate Radar service | Rate Radar | Daily cache |  |  |
| `public.persona_runs` | 45 | 2026-10-05 | persona layer (written per run) | Session 3 | Per persona run | Snapshot deck | Show last run on the command center |
| `public.rate_observations` | 867 | 2026-10-02 | Rate Radar service | Rate Radar | Per Rate Radar run | Intelligence Hub, Opportunity View, Rate Radar page | Show last run on the command center |
| `public.rate_radar_runs` | 46 |  | Rate Radar service | Rate Radar | Per Rate Radar run |  |  |
| `public.resonate_audience_enrichment` | 17 |  | Resonate integration (fetched) | Session 3 | On demand | Bank Assessment doc | Confirm owner and refresh trigger |
| `public.resonate_audiences` | 7 | 2026-10-02 | Resonate integration (fetched) | Session 3 | On demand | Bank Assessment doc | Confirm owner and refresh trigger |
| `public.url_pattern_registry` | 71 |  | Rate Radar service | Rate Radar | Per Rate Radar run |  |  |
| `raw.raw_rate_radar` | 324 | 2026-10-02 | Rate Radar service | Rate Radar | Per Rate Radar run | Intelligence Hub, Opportunity View, Rate Radar page | Show last run on the command center |

### APP (10)

| Object | Rows | Newest data | How it is refreshed | Owner | Cadence | Feeds | Action / note |
|---|---|---|---|---|---|---|---|
| `public.assessment_jobs` | 27 |  | written by the Hub assessment job queue (main.py) | App | - |  |  |
| `public.engagement_clauses` | 47 |  | written by the engagement tool | App | - |  |  |
| `public.engagement_drafts` | 0 |  | written by the engagement tool | App | - |  |  |
| `public.engagement_events` | 0 |  | written by the engagement tool | App | - |  |  |
| `public.engagement_module_text` | 56 |  | written by the engagement tool | App | - |  |  |
| `public.engagement_modules` | 5 |  | written by the engagement tool | App | - |  |  |
| `public.engagement_versions` | 0 |  | written by the engagement tool | App | - |  |  |
| `public.engagements` | 0 |  | written by the engagement tool | App | - |  |  |
| `public.pipeline_jobs` | 13 |  | written by the command center API | App | - |  |  |
| `public.profiles` | 2 |  | written by Supabase Auth + Users tab | App | - |  |  |

### DERIVED (21)

| Object | Rows | Newest data | How it is refreshed | Owner | Cadence | Feeds | Action / note |
|---|---|---|---|---|---|---|---|
| `geo.dim_zip_cbsa` | view |  | view | Session 1 | - |  | Base: geo.CBSA_zip |
| `pbi.vw_branch_competitors_10mi` | view |  | view (Power BI schema) | Session 1 | - |  | Power BI legacy; follows its base tables Base: pbi.branches_master |
| `pbi.vw_branch_portfolio_matrix` | view |  | view (Power BI schema) | Session 1 | - |  | Power BI legacy; follows its base tables Base: analytics.branch_opportunity_base, raw.raw_zhvi |
| `pbi.vw_branch_radius_stats_10mi` | view |  | view (Power BI schema) | Session 1 | - |  | Power BI legacy; follows its base tables Base: pbi.branches_master |
| `pbi.vw_cert_to_bank_name` | view |  | view (Power BI schema) | Session 1 | - |  | Power BI legacy; follows its base tables Base: raw.raw_sod |
| `pbi.vw_cu_cost_of_funds` | view |  | view (Power BI schema) | Session 1 | - |  | Power BI legacy; follows its base tables Base: raw.Raw_cu_fs220, raw.raw_cu_branches |
| `pbi.vw_rate_radar_history_v1_legacy` | view |  | view (Power BI schema) | Session 1 | - |  | Power BI legacy; follows its base tables Base: raw.raw_rate_radar |
| `pbi.vw_rate_radar_latest_v1_legacy` | view |  | view (Power BI schema) | Session 1 | - |  | Power BI legacy; follows its base tables Base: raw.raw_rate_radar |
| `pbi.vw_rate_radar_triggers` | view |  | view (Power BI schema) | Session 1 | - |  | Power BI legacy; follows its base tables Base: raw.raw_rate_radar |
| `pbi.vw_zhvi_yoy` | view |  | view (Power BI schema) | Session 1 | - |  | Power BI legacy; follows its base tables Base: analytics.branch_opportunity_base, raw.raw_zhvi |
| `public.vw_bank_directory` | view |  | view | Session 1 | - |  | Base: ref.bank_website, ref.dim_institutions |
| `public.vw_bank_financial_snapshot` | view |  | view | Session 1 | - |  | Base: raw.Raw_cu_fs220, raw.raw_UBPR, raw.raw_schedule_RC, raw.raw_schedule_RI |
| `public.vw_branch_opportunity_cbsa` | view |  | view | Session 1 | - |  | Base: analytics.branch_opportunity_base, geo.CBSA_zip, raw.raw_population |
| `public.vw_branch_opportunity_yoy_delta` | view |  | view | Session 1 | - |  | Base: analytics.branch_opportunity_base, analytics.branch_opportunity_base_history |
| `public.vw_cfpb_complaints_wow` | view |  | view | Session 1 | - |  | Base: raw.raw_cfpb_complaints, raw.raw_cfpb_complaints_trend |
| `public.vw_network_top_targets` | view |  | view | Session 1 | - |  | Base: public.network_top_targets |
| `public.vw_prospecting_score` | view |  | view | Session 1 | - |  | Base: analytics.bank_financial_snapshot_latest, raw.raw_sod, ref.dim_institutions |
| `public.vw_rate_radar_history` | view |  | view | Session 1 | - |  | Base: public.rate_observations, raw.raw_rate_radar |
| `public.vw_rate_radar_latest` | view |  | view | Session 1 | - |  | Base: public.rate_observations, raw.raw_rate_radar |
| `public.vw_smb_index_by_zip` | view |  | view | Session 1 | - |  | Base: raw.raw_business_formation_state, raw.raw_cbp_totals, raw.raw_population |
| `public.vw_zip_persona` | view |  | view | Session 1 | - |  | Base: raw.raw_income, raw.raw_occupation, raw.raw_population |

### SYSTEM (3)

| Object | Rows | Newest data | How it is refreshed | Owner | Cadence | Feeds | Action / note |
|---|---|---|---|---|---|---|---|
| `public.geography_columns` | view |  | PostGIS system view | - | - |  |  |
| `public.geometry_columns` | view |  | PostGIS system view | - | - |  |  |
| `public.spatial_ref_sys` | 8,500 |  | PostGIS system table | - | - |  | Security finding a92 (anon write grants, RLS off); Coordinator owns the fix; do not touch |

