-- a1 Test 2 (Phase 2) — exact queries run against analytics.bank_financial_snapshot,
-- per docs/a1_deposit_flight_backtest_methodology.md sections 2.2/2.3.
-- Run via Supabase MCP execute_sql against project tuiiywphoynbmkxpoyps on 2026-09-13.
-- Results captured in docs/a1_test2_results.md.
--
-- Template (repeated per predictor x transition): swap the predictor column in `base`,
-- and the period literals for C1 (2024-12-31 -> 2025-03-31) or C2 (2025-09-30 -> 2025-12-31).

-- ============================================================
-- ROA, Transition C1, WITH M&A exclusion
-- ============================================================
with base as (
  select inst_key, total_assets, total_deposits, roa as predictor_val
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2024-12-31'
),
outc as (
  select inst_key, total_deposits as dep_outc
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2025-03-31'
),
joined as (
  select b.inst_key, b.total_assets, b.predictor_val, b.total_deposits as dep_pred, o.dep_outc,
         greatest(least((o.dep_outc - b.total_deposits)/nullif(b.total_deposits,0), 0.20), -0.20) as dep_growth
  from base b join outc o on b.inst_key=o.inst_key
  where b.total_deposits >= 10000000
    and o.dep_outc >= 0.05 * b.total_deposits
    and b.predictor_val is not null
),
tercile as (
  select *, ntile(3) over (order by total_assets) as size_tercile,
            ntile(3) over (order by predictor_val) as pred_tercile
  from joined
)
select size_tercile, pred_tercile, count(*) n,
       round(avg(dep_growth)::numeric*100,2) avg_growth_pct,
       round(percentile_cont(0.5) within group(order by dep_growth)::numeric*100,2) median_growth_pct
from tercile group by 1,2 order by 1,2;

-- ============================================================
-- ROA, Transition C2, WITH M&A exclusion
-- ============================================================
with base as (
  select inst_key, total_assets, total_deposits, roa as predictor_val
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2025-09-30'
),
outc as (
  select inst_key, total_deposits as dep_outc
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2025-12-31'
),
joined as (
  select b.inst_key, b.total_assets, b.predictor_val, b.total_deposits as dep_pred, o.dep_outc,
         greatest(least((o.dep_outc - b.total_deposits)/nullif(b.total_deposits,0), 0.20), -0.20) as dep_growth
  from base b join outc o on b.inst_key=o.inst_key
  where b.total_deposits >= 10000000
    and o.dep_outc >= 0.05 * b.total_deposits
    and b.predictor_val is not null
),
tercile as (
  select *, ntile(3) over (order by total_assets) as size_tercile,
            ntile(3) over (order by predictor_val) as pred_tercile
  from joined
)
select size_tercile, pred_tercile, count(*) n,
       round(avg(dep_growth)::numeric*100,2) avg_growth_pct,
       round(percentile_cont(0.5) within group(order by dep_growth)::numeric*100,2) median_growth_pct
from tercile group by 1,2 order by 1,2;

-- ============================================================
-- LDR (loans_to_deposits_pct), Transition C1, WITH M&A exclusion
-- ============================================================
with base as (
  select inst_key, total_assets, total_deposits, loans_to_deposits_pct as predictor_val
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2024-12-31'
),
outc as (
  select inst_key, total_deposits as dep_outc
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2025-03-31'
),
joined as (
  select b.inst_key, b.total_assets, b.predictor_val, b.total_deposits as dep_pred, o.dep_outc,
         greatest(least((o.dep_outc - b.total_deposits)/nullif(b.total_deposits,0), 0.20), -0.20) as dep_growth
  from base b join outc o on b.inst_key=o.inst_key
  where b.total_deposits >= 10000000
    and o.dep_outc >= 0.05 * b.total_deposits
    and b.predictor_val is not null
),
tercile as (
  select *, ntile(3) over (order by total_assets) as size_tercile,
            ntile(3) over (order by predictor_val) as pred_tercile
  from joined
)
select size_tercile, pred_tercile, count(*) n,
       round(avg(dep_growth)::numeric*100,2) avg_growth_pct,
       round(percentile_cont(0.5) within group(order by dep_growth)::numeric*100,2) median_growth_pct
from tercile group by 1,2 order by 1,2;

-- ============================================================
-- LDR, Transition C2, WITH M&A exclusion
-- ============================================================
with base as (
  select inst_key, total_assets, total_deposits, loans_to_deposits_pct as predictor_val
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2025-09-30'
),
outc as (
  select inst_key, total_deposits as dep_outc
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2025-12-31'
),
joined as (
  select b.inst_key, b.total_assets, b.predictor_val, b.total_deposits as dep_pred, o.dep_outc,
         greatest(least((o.dep_outc - b.total_deposits)/nullif(b.total_deposits,0), 0.20), -0.20) as dep_growth
  from base b join outc o on b.inst_key=o.inst_key
  where b.total_deposits >= 10000000
    and o.dep_outc >= 0.05 * b.total_deposits
    and b.predictor_val is not null
),
tercile as (
  select *, ntile(3) over (order by total_assets) as size_tercile,
            ntile(3) over (order by predictor_val) as pred_tercile
  from joined
)
select size_tercile, pred_tercile, count(*) n,
       round(avg(dep_growth)::numeric*100,2) avg_growth_pct,
       round(percentile_cont(0.5) within group(order by dep_growth)::numeric*100,2) median_growth_pct
from tercile group by 1,2 order by 1,2;

-- ============================================================
-- Brokered Deposits %, Transition C1, WITH M&A exclusion
-- ============================================================
with base as (
  select inst_key, total_assets, total_deposits, brokered_deposits_pct as predictor_val
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2024-12-31'
),
outc as (
  select inst_key, total_deposits as dep_outc
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2025-03-31'
),
joined as (
  select b.inst_key, b.total_assets, b.predictor_val, b.total_deposits as dep_pred, o.dep_outc,
         greatest(least((o.dep_outc - b.total_deposits)/nullif(b.total_deposits,0), 0.20), -0.20) as dep_growth
  from base b join outc o on b.inst_key=o.inst_key
  where b.total_deposits >= 10000000
    and o.dep_outc >= 0.05 * b.total_deposits
    and b.predictor_val is not null
),
tercile as (
  select *, ntile(3) over (order by total_assets) as size_tercile,
            ntile(3) over (order by predictor_val) as pred_tercile
  from joined
)
select size_tercile, pred_tercile, count(*) n,
       round(avg(dep_growth)::numeric*100,2) avg_growth_pct,
       round(percentile_cont(0.5) within group(order by dep_growth)::numeric*100,2) median_growth_pct
from tercile group by 1,2 order by 1,2;

-- ============================================================
-- Brokered Deposits %, Transition C2, WITH M&A exclusion
-- ============================================================
with base as (
  select inst_key, total_assets, total_deposits, brokered_deposits_pct as predictor_val
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2025-09-30'
),
outc as (
  select inst_key, total_deposits as dep_outc
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2025-12-31'
),
joined as (
  select b.inst_key, b.total_assets, b.predictor_val, b.total_deposits as dep_pred, o.dep_outc,
         greatest(least((o.dep_outc - b.total_deposits)/nullif(b.total_deposits,0), 0.20), -0.20) as dep_growth
  from base b join outc o on b.inst_key=o.inst_key
  where b.total_deposits >= 10000000
    and o.dep_outc >= 0.05 * b.total_deposits
    and b.predictor_val is not null
),
tercile as (
  select *, ntile(3) over (order by total_assets) as size_tercile,
            ntile(3) over (order by predictor_val) as pred_tercile
  from joined
)
select size_tercile, pred_tercile, count(*) n,
       round(avg(dep_growth)::numeric*100,2) avg_growth_pct,
       round(percentile_cont(0.5) within group(order by dep_growth)::numeric*100,2) median_growth_pct
from tercile group by 1,2 order by 1,2;

-- ============================================================
-- NIM, Transition C1, WITH M&A exclusion
-- ============================================================
with base as (
  select inst_key, total_assets, total_deposits, nim as predictor_val
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2024-12-31'
),
outc as (
  select inst_key, total_deposits as dep_outc
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2025-03-31'
),
joined as (
  select b.inst_key, b.total_assets, b.predictor_val, b.total_deposits as dep_pred, o.dep_outc,
         greatest(least((o.dep_outc - b.total_deposits)/nullif(b.total_deposits,0), 0.20), -0.20) as dep_growth
  from base b join outc o on b.inst_key=o.inst_key
  where b.total_deposits >= 10000000
    and o.dep_outc >= 0.05 * b.total_deposits
    and b.predictor_val is not null
),
tercile as (
  select *, ntile(3) over (order by total_assets) as size_tercile,
            ntile(3) over (order by predictor_val) as pred_tercile
  from joined
)
select size_tercile, pred_tercile, count(*) n,
       round(avg(dep_growth)::numeric*100,2) avg_growth_pct,
       round(percentile_cont(0.5) within group(order by dep_growth)::numeric*100,2) median_growth_pct
from tercile group by 1,2 order by 1,2;

-- ============================================================
-- NIM, Transition C1, WITHOUT M&A exclusion (criterion 3 comparison)
-- ============================================================
with base as (
  select inst_key, total_assets, total_deposits, nim as predictor_val
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2024-12-31'
),
outc as (
  select inst_key, total_deposits as dep_outc
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2025-03-31'
),
joined as (
  select b.inst_key, b.total_assets, b.predictor_val, b.total_deposits as dep_pred, o.dep_outc,
         greatest(least((o.dep_outc - b.total_deposits)/nullif(b.total_deposits,0), 0.20), -0.20) as dep_growth
  from base b join outc o on b.inst_key=o.inst_key
  where b.total_deposits >= 10000000
    and b.predictor_val is not null
),
tercile as (
  select *, ntile(3) over (order by total_assets) as size_tercile,
            ntile(3) over (order by predictor_val) as pred_tercile
  from joined
)
select size_tercile, pred_tercile, count(*) n,
       round(percentile_cont(0.5) within group(order by dep_growth)::numeric*100,2) median_growth_pct
from tercile group by 1,2 order by 1,2;

-- ============================================================
-- NIM, Transition C2, WITH M&A exclusion
-- ============================================================
with base as (
  select inst_key, total_assets, total_deposits, nim as predictor_val
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2025-09-30'
),
outc as (
  select inst_key, total_deposits as dep_outc
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2025-12-31'
),
joined as (
  select b.inst_key, b.total_assets, b.predictor_val, b.total_deposits as dep_pred, o.dep_outc,
         greatest(least((o.dep_outc - b.total_deposits)/nullif(b.total_deposits,0), 0.20), -0.20) as dep_growth
  from base b join outc o on b.inst_key=o.inst_key
  where b.total_deposits >= 10000000
    and o.dep_outc >= 0.05 * b.total_deposits
    and b.predictor_val is not null
),
tercile as (
  select *, ntile(3) over (order by total_assets) as size_tercile,
            ntile(3) over (order by predictor_val) as pred_tercile
  from joined
)
select size_tercile, pred_tercile, count(*) n,
       round(avg(dep_growth)::numeric*100,2) avg_growth_pct,
       round(percentile_cont(0.5) within group(order by dep_growth)::numeric*100,2) median_growth_pct
from tercile group by 1,2 order by 1,2;
