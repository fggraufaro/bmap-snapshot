-- a1 Test 4 (Phase 2) — local market tailwind (income/population/ZHVI) -> branch deposit growth.
-- Per docs/a1_deposit_flight_backtest_methodology.md section 4.2. Run 2026-09-13.
-- Results in docs/a1_test4_results.md.

-- ============================================================
-- Income, Transition A: growth 2021->2022 predicts branch SOD growth 2023->2024
-- ============================================================
with income_zip as (
  select right("Geographic Area Name",5) as zip, "YEAR",
         nullif(regexp_replace(income, '[^0-9.]', '', 'g'), '')::numeric as income
  from raw.raw_income where "YEAR" in ('2021','2022')
),
income_growth as (
  select a.zip, (b.income - a.income)/nullif(a.income,0) as income_growth
  from income_zip a join income_zip b on a.zip=b.zip and a."YEAR"='2021' and b."YEAR"='2022'
  where a.income > 0 and b.income is not null
),
dedup as (
  select distinct "YEAR","UNINUMBR","STCNTYBR","DEPSUMBR","ZIPBR"
  from raw.raw_sod where "YEAR" in ('2023','2024') and "DEPSUMBR" is not null and "STCNTYBR" is not null
),
county_totals as (
  select "YEAR","STCNTYBR", sum("DEPSUMBR") as county_dep from dedup group by 1,2
),
branch_pred as (
  select d."UNINUMBR", d."ZIPBR", d."DEPSUMBR" as dep_pred,
         d."DEPSUMBR"::numeric / nullif(ct.county_dep,0) as own_share
  from dedup d join county_totals ct on ct."YEAR"=d."YEAR" and ct."STCNTYBR"=d."STCNTYBR"
  where d."YEAR"='2023'
),
branch_outc as (select "UNINUMBR", "DEPSUMBR" as dep_outc from dedup where "YEAR"='2024'),
joined as (
  select bp.*, bo.dep_outc, ig.income_growth,
         greatest(least((bo.dep_outc::numeric - bp.dep_pred::numeric)/nullif(bp.dep_pred,0), 2.0), -0.9) as yoy_growth
  from branch_pred bp
  join branch_outc bo on bp."UNINUMBR"=bo."UNINUMBR"
  join income_growth ig on ig.zip = bp."ZIPBR"
  where bp.dep_pred >= 1000 and bo.dep_outc::numeric >= 0.05 * bp.dep_pred
),
tercile as (
  select *, ntile(3) over (order by own_share) as share_tercile,
            ntile(3) over (order by income_growth) as pred_tercile
  from joined
)
select share_tercile, pred_tercile, count(*) n,
       round(percentile_cont(0.5) within group(order by yoy_growth)::numeric*100,2) median_growth_pct
from tercile group by 1,2 order by 1,2;

-- ============================================================
-- Income, Transition B: growth 2022->2023 predicts branch SOD growth 2024->2025
-- ============================================================
with income_zip as (
  select right("Geographic Area Name",5) as zip, "YEAR",
         nullif(regexp_replace(income, '[^0-9.]', '', 'g'), '')::numeric as income
  from raw.raw_income where "YEAR" in ('2022','2023')
),
income_growth as (
  select a.zip, (b.income - a.income)/nullif(a.income,0) as income_growth
  from income_zip a join income_zip b on a.zip=b.zip and a."YEAR"='2022' and b."YEAR"='2023'
  where a.income > 0 and b.income is not null
),
dedup as (
  select distinct "YEAR","UNINUMBR","STCNTYBR","DEPSUMBR","ZIPBR"
  from raw.raw_sod where "YEAR" in ('2024','2025') and "DEPSUMBR" is not null and "STCNTYBR" is not null
),
county_totals as (
  select "YEAR","STCNTYBR", sum("DEPSUMBR") as county_dep from dedup group by 1,2
),
branch_pred as (
  select d."UNINUMBR", d."ZIPBR", d."DEPSUMBR" as dep_pred,
         d."DEPSUMBR"::numeric / nullif(ct.county_dep,0) as own_share
  from dedup d join county_totals ct on ct."YEAR"=d."YEAR" and ct."STCNTYBR"=d."STCNTYBR"
  where d."YEAR"='2024'
),
branch_outc as (select "UNINUMBR", "DEPSUMBR" as dep_outc from dedup where "YEAR"='2025'),
joined as (
  select bp.*, bo.dep_outc, ig.income_growth,
         greatest(least((bo.dep_outc::numeric - bp.dep_pred::numeric)/nullif(bp.dep_pred,0), 2.0), -0.9) as yoy_growth
  from branch_pred bp
  join branch_outc bo on bp."UNINUMBR"=bo."UNINUMBR"
  join income_growth ig on ig.zip = bp."ZIPBR"
  where bp.dep_pred >= 1000 and bo.dep_outc::numeric >= 0.05 * bp.dep_pred
),
tercile as (
  select *, ntile(3) over (order by own_share) as share_tercile,
            ntile(3) over (order by income_growth) as pred_tercile
  from joined
)
select share_tercile, pred_tercile, count(*) n,
       round(percentile_cont(0.5) within group(order by yoy_growth)::numeric*100,2) median_growth_pct
from tercile group by 1,2 order by 1,2;

-- ============================================================
-- Population, Transition A: growth 2021->2022 predicts branch SOD growth 2023->2024
-- ============================================================
with pop_zip as (
  select right("Geographic Area Name",5) as zip, "YEAR", total::numeric as total
  from raw.raw_population where "YEAR" in ('2021','2022')
),
pop_growth as (
  select a.zip, (b.total - a.total)/nullif(a.total,0) as pop_growth
  from pop_zip a join pop_zip b on a.zip=b.zip and a."YEAR"='2021' and b."YEAR"='2022'
  where a.total > 0 and b.total is not null
),
dedup as (
  select distinct "YEAR","UNINUMBR","STCNTYBR","DEPSUMBR","ZIPBR"
  from raw.raw_sod where "YEAR" in ('2023','2024') and "DEPSUMBR" is not null and "STCNTYBR" is not null
),
county_totals as (
  select "YEAR","STCNTYBR", sum("DEPSUMBR") as county_dep from dedup group by 1,2
),
branch_pred as (
  select d."UNINUMBR", d."ZIPBR", d."DEPSUMBR" as dep_pred,
         d."DEPSUMBR"::numeric / nullif(ct.county_dep,0) as own_share
  from dedup d join county_totals ct on ct."YEAR"=d."YEAR" and ct."STCNTYBR"=d."STCNTYBR"
  where d."YEAR"='2023'
),
branch_outc as (select "UNINUMBR", "DEPSUMBR" as dep_outc from dedup where "YEAR"='2024'),
joined as (
  select bp.*, bo.dep_outc, pg.pop_growth,
         greatest(least((bo.dep_outc::numeric - bp.dep_pred::numeric)/nullif(bp.dep_pred,0), 2.0), -0.9) as yoy_growth
  from branch_pred bp
  join branch_outc bo on bp."UNINUMBR"=bo."UNINUMBR"
  join pop_growth pg on pg.zip = bp."ZIPBR"
  where bp.dep_pred >= 1000 and bo.dep_outc::numeric >= 0.05 * bp.dep_pred
),
tercile as (
  select *, ntile(3) over (order by own_share) as share_tercile,
            ntile(3) over (order by pop_growth) as pred_tercile
  from joined
)
select share_tercile, pred_tercile, count(*) n,
       round(percentile_cont(0.5) within group(order by yoy_growth)::numeric*100,2) median_growth_pct
from tercile group by 1,2 order by 1,2;

-- ============================================================
-- Population, Transition B: growth 2022->2023 predicts branch SOD growth 2024->2025
-- ============================================================
with pop_zip as (
  select right("Geographic Area Name",5) as zip, "YEAR", total::numeric as total
  from raw.raw_population where "YEAR" in ('2022','2023')
),
pop_growth as (
  select a.zip, (b.total - a.total)/nullif(a.total,0) as pop_growth
  from pop_zip a join pop_zip b on a.zip=b.zip and a."YEAR"='2022' and b."YEAR"='2023'
  where a.total > 0 and b.total is not null
),
dedup as (
  select distinct "YEAR","UNINUMBR","STCNTYBR","DEPSUMBR","ZIPBR"
  from raw.raw_sod where "YEAR" in ('2024','2025') and "DEPSUMBR" is not null and "STCNTYBR" is not null
),
county_totals as (
  select "YEAR","STCNTYBR", sum("DEPSUMBR") as county_dep from dedup group by 1,2
),
branch_pred as (
  select d."UNINUMBR", d."ZIPBR", d."DEPSUMBR" as dep_pred,
         d."DEPSUMBR"::numeric / nullif(ct.county_dep,0) as own_share
  from dedup d join county_totals ct on ct."YEAR"=d."YEAR" and ct."STCNTYBR"=d."STCNTYBR"
  where d."YEAR"='2024'
),
branch_outc as (select "UNINUMBR", "DEPSUMBR" as dep_outc from dedup where "YEAR"='2025'),
joined as (
  select bp.*, bo.dep_outc, pg.pop_growth,
         greatest(least((bo.dep_outc::numeric - bp.dep_pred::numeric)/nullif(bp.dep_pred,0), 2.0), -0.9) as yoy_growth
  from branch_pred bp
  join branch_outc bo on bp."UNINUMBR"=bo."UNINUMBR"
  join pop_growth pg on pg.zip = bp."ZIPBR"
  where bp.dep_pred >= 1000 and bo.dep_outc::numeric >= 0.05 * bp.dep_pred
),
tercile as (
  select *, ntile(3) over (order by own_share) as share_tercile,
            ntile(3) over (order by pop_growth) as pred_tercile
  from joined
)
select share_tercile, pred_tercile, count(*) n,
       round(percentile_cont(0.5) within group(order by yoy_growth)::numeric*100,2) median_growth_pct
from tercile group by 1,2 order by 1,2;

-- ============================================================
-- ZHVI, single transition: growth Jan2023->Jan2024 predicts branch SOD growth 2024->2025
-- (simplified vs. the raw_sod dedup pattern used above -- 2024/2025 SOD confirmed dup-free,
-- so the DISTINCT step is skipped here purely for query performance; the original version
-- with DISTINCT timed out on this larger join)
-- ============================================================
with zhvi_growth as (
  select a.zip, (b.zhvi::numeric - a.zhvi::numeric)/nullif(a.zhvi::numeric,0) as zhvi_growth
  from raw.raw_zhvi a join raw.raw_zhvi b on a.zip=b.zip
  where a.period='2023-01-31' and b.period='2024-01-31' and a.zhvi::numeric > 0
),
county_totals as (
  select "STCNTYBR", sum("DEPSUMBR") as county_dep
  from raw.raw_sod where "YEAR"='2024' and "DEPSUMBR" is not null and "STCNTYBR" is not null
  group by 1
),
branch_pred as (
  select d."UNINUMBR", d."ZIPBR", d."DEPSUMBR" as dep_pred,
         d."DEPSUMBR"::numeric / nullif(ct.county_dep,0) as own_share
  from raw.raw_sod d join county_totals ct on ct."STCNTYBR"=d."STCNTYBR"
  where d."YEAR"='2024' and d."DEPSUMBR" is not null
),
branch_outc as (select "UNINUMBR", "DEPSUMBR" as dep_outc from raw.raw_sod where "YEAR"='2025' and "DEPSUMBR" is not null),
joined as (
  select bp.*, bo.dep_outc, zg.zhvi_growth,
         greatest(least((bo.dep_outc::numeric - bp.dep_pred::numeric)/nullif(bp.dep_pred,0), 2.0), -0.9) as yoy_growth
  from branch_pred bp
  join branch_outc bo on bp."UNINUMBR"=bo."UNINUMBR"
  join zhvi_growth zg on zg.zip = bp."ZIPBR"
  where bp.dep_pred >= 1000 and bo.dep_outc::numeric >= 0.05 * bp.dep_pred
),
tercile as (
  select *, ntile(3) over (order by own_share) as share_tercile,
            ntile(3) over (order by zhvi_growth) as pred_tercile
  from joined
)
select share_tercile, pred_tercile, count(*) n,
       round(percentile_cont(0.5) within group(order by yoy_growth)::numeric*100,2) median_growth_pct
from tercile group by 1,2 order by 1,2;
