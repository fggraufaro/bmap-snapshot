-- a1 Test 1 (Phase 2) — exact queries run against raw.raw_sod, per
-- docs/a1_deposit_flight_backtest_methodology.md section 1.2.
-- Run via Supabase MCP execute_sql against project tuiiywphoynbmkxpoyps on 2026-09-13.
-- Results captured in docs/a1_test1_results.md.

-- ============================================================
-- Transition A: 2023 (predictor) -> 2024 (outcome), WITH closure exclusion
-- ============================================================
with dedup as (
  select distinct "YEAR","UNINUMBR","RSSDID","STCNTYBR","DEPSUMBR"
  from raw.raw_sod
  where "YEAR" in ('2023','2024') and "DEPSUMBR" is not null and "STCNTYBR" is not null
),
county_totals as (
  select "YEAR","STCNTYBR", sum("DEPSUMBR") as county_dep, count(*) as county_branches
  from dedup group by 1,2
),
pred as (
  select d."UNINUMBR", d."DEPSUMBR" as dep_pred,
         d."DEPSUMBR"::numeric / nullif(ct.county_dep,0) as own_share,
         (ct.county_branches - 1) as competitor_count
  from dedup d join county_totals ct on ct."YEAR"=d."YEAR" and ct."STCNTYBR"=d."STCNTYBR"
  where d."YEAR"='2023'
),
outc as (select "UNINUMBR", "DEPSUMBR" as dep_outc from dedup where "YEAR"='2024'),
joined as (
  select a.*, b.dep_outc,
         greatest(least((b.dep_outc::numeric - a.dep_pred::numeric)/nullif(a.dep_pred,0), 2.0), -0.9) as yoy_growth
  from pred a join outc b on a."UNINUMBR"=b."UNINUMBR"
  where a.dep_pred >= 1000 and b.dep_outc::numeric >= 0.05 * a.dep_pred
),
tercile as (
  select *, ntile(3) over (order by own_share) as share_tercile,
            ntile(3) over (order by competitor_count) as comp_tercile
  from joined
)
select share_tercile, comp_tercile, count(*) as n,
       round(avg(yoy_growth)::numeric*100,2) as avg_growth_pct,
       round(percentile_cont(0.5) within group (order by yoy_growth)::numeric*100,2) as median_growth_pct
from tercile group by 1,2 order by 1,2;

-- ============================================================
-- Transition A: 2023 -> 2024, WITHOUT closure exclusion (criterion 3 comparison)
-- ============================================================
with dedup as (
  select distinct "YEAR","UNINUMBR","RSSDID","STCNTYBR","DEPSUMBR"
  from raw.raw_sod
  where "YEAR" in ('2023','2024') and "DEPSUMBR" is not null and "STCNTYBR" is not null
),
county_totals as (
  select "YEAR","STCNTYBR", sum("DEPSUMBR") as county_dep, count(*) as county_branches
  from dedup group by 1,2
),
pred as (
  select d."UNINUMBR", d."DEPSUMBR" as dep_pred,
         d."DEPSUMBR"::numeric / nullif(ct.county_dep,0) as own_share,
         (ct.county_branches - 1) as competitor_count
  from dedup d join county_totals ct on ct."YEAR"=d."YEAR" and ct."STCNTYBR"=d."STCNTYBR"
  where d."YEAR"='2023'
),
outc as (select "UNINUMBR", "DEPSUMBR" as dep_outc from dedup where "YEAR"='2024'),
joined as (
  select a.*, b.dep_outc,
         greatest(least((b.dep_outc::numeric - a.dep_pred::numeric)/nullif(a.dep_pred,0), 2.0), -0.9) as yoy_growth
  from pred a join outc b on a."UNINUMBR"=b."UNINUMBR"
  where a.dep_pred >= 1000   -- no 5%-closure floor
),
tercile as (
  select *, ntile(3) over (order by own_share) as share_tercile,
            ntile(3) over (order by competitor_count) as comp_tercile
  from joined
)
select share_tercile, comp_tercile, count(*) as n,
       round(percentile_cont(0.5) within group (order by yoy_growth)::numeric*100,2) as median_growth_pct
from tercile group by 1,2 order by 1,2;

-- ============================================================
-- Transition B: 2024 (predictor) -> 2025 (outcome), WITH closure exclusion
-- ============================================================
with dedup as (
  select distinct "YEAR","UNINUMBR","RSSDID","STCNTYBR","DEPSUMBR"
  from raw.raw_sod
  where "YEAR" in ('2024','2025') and "DEPSUMBR" is not null and "STCNTYBR" is not null
),
county_totals as (
  select "YEAR","STCNTYBR", sum("DEPSUMBR") as county_dep, count(*) as county_branches
  from dedup group by 1,2
),
pred as (
  select d."UNINUMBR", d."DEPSUMBR" as dep_pred,
         d."DEPSUMBR"::numeric / nullif(ct.county_dep,0) as own_share,
         (ct.county_branches - 1) as competitor_count
  from dedup d join county_totals ct on ct."YEAR"=d."YEAR" and ct."STCNTYBR"=d."STCNTYBR"
  where d."YEAR"='2024'
),
outc as (select "UNINUMBR", "DEPSUMBR" as dep_outc from dedup where "YEAR"='2025'),
joined as (
  select a.*, b.dep_outc,
         greatest(least((b.dep_outc::numeric - a.dep_pred::numeric)/nullif(a.dep_pred,0), 2.0), -0.9) as yoy_growth
  from pred a join outc b on a."UNINUMBR"=b."UNINUMBR"
  where a.dep_pred >= 1000 and b.dep_outc::numeric >= 0.05 * a.dep_pred
),
tercile as (
  select *, ntile(3) over (order by own_share) as share_tercile,
            ntile(3) over (order by competitor_count) as comp_tercile
  from joined
)
select share_tercile, comp_tercile, count(*) as n,
       round(avg(yoy_growth)::numeric*100,2) as avg_growth_pct,
       round(percentile_cont(0.5) within group (order by yoy_growth)::numeric*100,2) as median_growth_pct
from tercile group by 1,2 order by 1,2;

-- ============================================================
-- Transition B: 2024 -> 2025, WITHOUT closure exclusion (criterion 3 comparison)
-- ============================================================
with dedup as (
  select distinct "YEAR","UNINUMBR","RSSDID","STCNTYBR","DEPSUMBR"
  from raw.raw_sod
  where "YEAR" in ('2024','2025') and "DEPSUMBR" is not null and "STCNTYBR" is not null
),
county_totals as (
  select "YEAR","STCNTYBR", sum("DEPSUMBR") as county_dep, count(*) as county_branches
  from dedup group by 1,2
),
pred as (
  select d."UNINUMBR", d."DEPSUMBR" as dep_pred,
         d."DEPSUMBR"::numeric / nullif(ct.county_dep,0) as own_share,
         (ct.county_branches - 1) as competitor_count
  from dedup d join county_totals ct on ct."YEAR"=d."YEAR" and ct."STCNTYBR"=d."STCNTYBR"
  where d."YEAR"='2024'
),
outc as (select "UNINUMBR", "DEPSUMBR" as dep_outc from dedup where "YEAR"='2025'),
joined as (
  select a.*, b.dep_outc,
         greatest(least((b.dep_outc::numeric - a.dep_pred::numeric)/nullif(a.dep_pred,0), 2.0), -0.9) as yoy_growth
  from pred a join outc b on a."UNINUMBR"=b."UNINUMBR"
  where a.dep_pred >= 1000
),
tercile as (
  select *, ntile(3) over (order by own_share) as share_tercile,
            ntile(3) over (order by competitor_count) as comp_tercile
  from joined
)
select share_tercile, comp_tercile, count(*) as n,
       round(percentile_cont(0.5) within group (order by yoy_growth)::numeric*100,2) as median_growth_pct
from tercile group by 1,2 order by 1,2;
