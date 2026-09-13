-- a1 Test 3 (Phase 2, directional only) — parent-institution Call Report financials
-- (2024-12-31) predicting one specific branch's 2024->2025 SOD deposit growth.
-- Per docs/a1_deposit_flight_backtest_methodology.md section 3.2. Single transition only
-- (see 3.2's documented limitation) -- results in docs/a1_test3_results.md are directional,
-- no pass/fail verdict.

-- ============================================================
-- ROA
-- ============================================================
with dedup as (
  select distinct "YEAR","UNINUMBR","RSSDID","STCNTYBR","DEPSUMBR"
  from raw.raw_sod where "YEAR" in ('2024','2025') and "DEPSUMBR" is not null and "STCNTYBR" is not null
),
county_totals as (
  select "YEAR","STCNTYBR", sum("DEPSUMBR") as county_dep from dedup group by 1,2
),
branch_pred as (
  select d."UNINUMBR", d."RSSDID", d."DEPSUMBR" as dep_pred,
         d."DEPSUMBR"::numeric / nullif(ct.county_dep,0) as own_share
  from dedup d join county_totals ct on ct."YEAR"=d."YEAR" and ct."STCNTYBR"=d."STCNTYBR"
  where d."YEAR"='2024'
),
branch_outc as (select "UNINUMBR", "DEPSUMBR" as dep_outc from dedup where "YEAR"='2025'),
parent as (
  select institution_id as rssdid, roa as predictor_val
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2024-12-31' and roa is not null
),
joined as (
  select bp.*, bo.dep_outc, p.predictor_val,
         greatest(least((bo.dep_outc::numeric - bp.dep_pred::numeric)/nullif(bp.dep_pred,0), 2.0), -0.9) as yoy_growth
  from branch_pred bp
  join branch_outc bo on bp."UNINUMBR"=bo."UNINUMBR"
  join parent p on p.rssdid = bp."RSSDID"
  where bp.dep_pred >= 1000 and bo.dep_outc::numeric >= 0.05 * bp.dep_pred
),
tercile as (
  select *, ntile(3) over (order by own_share) as share_tercile,
            ntile(3) over (order by predictor_val) as pred_tercile
  from joined
)
select share_tercile, pred_tercile, count(*) n,
       round(percentile_cont(0.5) within group(order by yoy_growth)::numeric*100,2) median_growth_pct
from tercile group by 1,2 order by 1,2;

-- ============================================================
-- LDR (loans_to_deposits_pct)
-- ============================================================
with dedup as (
  select distinct "YEAR","UNINUMBR","RSSDID","STCNTYBR","DEPSUMBR"
  from raw.raw_sod where "YEAR" in ('2024','2025') and "DEPSUMBR" is not null and "STCNTYBR" is not null
),
county_totals as (
  select "YEAR","STCNTYBR", sum("DEPSUMBR") as county_dep from dedup group by 1,2
),
branch_pred as (
  select d."UNINUMBR", d."RSSDID", d."DEPSUMBR" as dep_pred,
         d."DEPSUMBR"::numeric / nullif(ct.county_dep,0) as own_share
  from dedup d join county_totals ct on ct."YEAR"=d."YEAR" and ct."STCNTYBR"=d."STCNTYBR"
  where d."YEAR"='2024'
),
branch_outc as (select "UNINUMBR", "DEPSUMBR" as dep_outc from dedup where "YEAR"='2025'),
parent as (
  select institution_id as rssdid, loans_to_deposits_pct as predictor_val
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2024-12-31' and loans_to_deposits_pct is not null
),
joined as (
  select bp.*, bo.dep_outc, p.predictor_val,
         greatest(least((bo.dep_outc::numeric - bp.dep_pred::numeric)/nullif(bp.dep_pred,0), 2.0), -0.9) as yoy_growth
  from branch_pred bp
  join branch_outc bo on bp."UNINUMBR"=bo."UNINUMBR"
  join parent p on p.rssdid = bp."RSSDID"
  where bp.dep_pred >= 1000 and bo.dep_outc::numeric >= 0.05 * bp.dep_pred
),
tercile as (
  select *, ntile(3) over (order by own_share) as share_tercile,
            ntile(3) over (order by predictor_val) as pred_tercile
  from joined
)
select share_tercile, pred_tercile, count(*) n,
       round(percentile_cont(0.5) within group(order by yoy_growth)::numeric*100,2) median_growth_pct
from tercile group by 1,2 order by 1,2;

-- ============================================================
-- Brokered Deposits %
-- ============================================================
with dedup as (
  select distinct "YEAR","UNINUMBR","RSSDID","STCNTYBR","DEPSUMBR"
  from raw.raw_sod where "YEAR" in ('2024','2025') and "DEPSUMBR" is not null and "STCNTYBR" is not null
),
county_totals as (
  select "YEAR","STCNTYBR", sum("DEPSUMBR") as county_dep from dedup group by 1,2
),
branch_pred as (
  select d."UNINUMBR", d."RSSDID", d."DEPSUMBR" as dep_pred,
         d."DEPSUMBR"::numeric / nullif(ct.county_dep,0) as own_share
  from dedup d join county_totals ct on ct."YEAR"=d."YEAR" and ct."STCNTYBR"=d."STCNTYBR"
  where d."YEAR"='2024'
),
branch_outc as (select "UNINUMBR", "DEPSUMBR" as dep_outc from dedup where "YEAR"='2025'),
parent as (
  select institution_id as rssdid, brokered_deposits_pct as predictor_val
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2024-12-31' and brokered_deposits_pct is not null
),
joined as (
  select bp.*, bo.dep_outc, p.predictor_val,
         greatest(least((bo.dep_outc::numeric - bp.dep_pred::numeric)/nullif(bp.dep_pred,0), 2.0), -0.9) as yoy_growth
  from branch_pred bp
  join branch_outc bo on bp."UNINUMBR"=bo."UNINUMBR"
  join parent p on p.rssdid = bp."RSSDID"
  where bp.dep_pred >= 1000 and bo.dep_outc::numeric >= 0.05 * bp.dep_pred
),
tercile as (
  select *, ntile(3) over (order by own_share) as share_tercile,
            ntile(3) over (order by predictor_val) as pred_tercile
  from joined
)
select share_tercile, pred_tercile, count(*) n,
       round(percentile_cont(0.5) within group(order by yoy_growth)::numeric*100,2) median_growth_pct
from tercile group by 1,2 order by 1,2;

-- ============================================================
-- NIM
-- ============================================================
with dedup as (
  select distinct "YEAR","UNINUMBR","RSSDID","STCNTYBR","DEPSUMBR"
  from raw.raw_sod where "YEAR" in ('2024','2025') and "DEPSUMBR" is not null and "STCNTYBR" is not null
),
county_totals as (
  select "YEAR","STCNTYBR", sum("DEPSUMBR") as county_dep from dedup group by 1,2
),
branch_pred as (
  select d."UNINUMBR", d."RSSDID", d."DEPSUMBR" as dep_pred,
         d."DEPSUMBR"::numeric / nullif(ct.county_dep,0) as own_share
  from dedup d join county_totals ct on ct."YEAR"=d."YEAR" and ct."STCNTYBR"=d."STCNTYBR"
  where d."YEAR"='2024'
),
branch_outc as (select "UNINUMBR", "DEPSUMBR" as dep_outc from dedup where "YEAR"='2025'),
parent as (
  select institution_id as rssdid, nim as predictor_val
  from analytics.bank_financial_snapshot
  where institution_type='bank' and period='2024-12-31' and nim is not null
),
joined as (
  select bp.*, bo.dep_outc, p.predictor_val,
         greatest(least((bo.dep_outc::numeric - bp.dep_pred::numeric)/nullif(bp.dep_pred,0), 2.0), -0.9) as yoy_growth
  from branch_pred bp
  join branch_outc bo on bp."UNINUMBR"=bo."UNINUMBR"
  join parent p on p.rssdid = bp."RSSDID"
  where bp.dep_pred >= 1000 and bo.dep_outc::numeric >= 0.05 * bp.dep_pred
),
tercile as (
  select *, ntile(3) over (order by own_share) as share_tercile,
            ntile(3) over (order by predictor_val) as pred_tercile
  from joined
)
select share_tercile, pred_tercile, count(*) n,
       round(percentile_cont(0.5) within group(order by yoy_growth)::numeric*100,2) median_growth_pct
from tercile group by 1,2 order by 1,2;
