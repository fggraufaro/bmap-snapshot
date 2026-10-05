-- Mirrored copy of public.vw_bank_financial_snapshot, for git-tracked
-- readability and code review. THE DATABASE COPY IS AUTHORITATIVE -- this file
-- is a snapshot taken 2026-10-05 straight from
--   select pg_get_viewdef('public.vw_bank_financial_snapshot'::regclass, true)
-- (written by a script, not retyped). Re-fetch it after any change to the view.
--
-- Since the 2026-10 cost-of-funds fix (a84/a85, Session 3): for banks,
-- cost_of_funds_pct, save_interest_pct, cd_interest_pct and
-- noninterest_income_pct are trailing 12 months (current YTD + prior-year Dec -
-- prior-year same-period YTD; NULL when the prior-year inputs aren't loaded);
-- credit unions are annualized by 12 / month-of-period. The pre-fix definition is
-- in backup.bak_a84_cof_view_def.
-- Refreshing analytics.bank_financial_snapshot(_latest) from this view takes
-- minutes: stage, validate, back up, then swap (see Session 3's a84 pipeline spec).
--
CREATE OR REPLACE VIEW public.vw_bank_financial_snapshot AS
 WITH bank_base AS (
         SELECT 'bank_'::text || ri."IDRSSD" AS inst_key,
            ri."IDRSSD" AS institution_id,
            'bank'::text AS institution_type,
            ri.period,
            NULLIF(rc."RCON2170", ''::text)::numeric * 1000::numeric AS total_assets,
            NULLIF(rc."RCON2200", ''::text)::numeric * 1000::numeric AS total_deposits,
            NULLIF(rc."RCON3210", ''::text)::numeric * 1000::numeric AS total_equity,
            NULLIF(ri."RIAD4074", ''::text)::numeric * 1000::numeric AS net_interest_income,
            NULLIF(ri."RIAD4079", ''::text)::numeric * 1000::numeric AS noninterest_income,
            NULLIF(ri."RIAD4073", ''::text)::numeric * 1000::numeric AS total_interest_expense,
            NULLIF(ri."RIAD4093", ''::text)::numeric * 1000::numeric AS noninterest_expense,
            NULLIF(ri."RIAD4340", ''::text)::numeric * 1000::numeric AS net_income,
            round(
                CASE
                    WHEN EXTRACT(month FROM ri.period) = 12::numeric THEN NULLIF(ri."RIAD0093", ''::text)::numeric
                    WHEN NULLIF(rp."RIAD0093", ''::text) IS NOT NULL AND NULLIF(rd."RIAD0093", ''::text) IS NOT NULL THEN NULLIF(ri."RIAD0093", ''::text)::numeric + NULLIF(rd."RIAD0093", ''::text)::numeric - NULLIF(rp."RIAD0093", ''::text)::numeric
                    ELSE NULL::numeric
                END / NULLIF(NULLIF(rc."RCON2200", ''::text)::numeric, 0::numeric) * 100::numeric, 4) AS save_interest_pct,
            round(
                CASE
                    WHEN EXTRACT(month FROM ri.period) = 12::numeric THEN NULLIF(ri."RIADHK03", ''::text)::numeric
                    WHEN NULLIF(rp."RIADHK03", ''::text) IS NOT NULL AND NULLIF(rd."RIADHK03", ''::text) IS NOT NULL THEN NULLIF(ri."RIADHK03", ''::text)::numeric + NULLIF(rd."RIADHK03", ''::text)::numeric - NULLIF(rp."RIADHK03", ''::text)::numeric
                    ELSE NULL::numeric
                END / NULLIF(NULLIF(rc."RCON2200", ''::text)::numeric, 0::numeric) * 100::numeric, 4) AS cd_interest_pct,
            round(
                CASE
                    WHEN EXTRACT(month FROM ri.period) = 12::numeric THEN NULLIF(ri."RIAD4073", ''::text)::numeric
                    WHEN NULLIF(rp."RIAD4073", ''::text) IS NOT NULL AND NULLIF(rd."RIAD4073", ''::text) IS NOT NULL THEN NULLIF(ri."RIAD4073", ''::text)::numeric + NULLIF(rd."RIAD4073", ''::text)::numeric - NULLIF(rp."RIAD4073", ''::text)::numeric
                    ELSE NULL::numeric
                END / NULLIF(NULLIF(rc."RCON2200", ''::text)::numeric, 0::numeric) * 100::numeric, 4) AS cost_of_funds_pct,
            round(
                CASE
                    WHEN EXTRACT(month FROM ri.period) = 12::numeric THEN NULLIF(ri."RIAD4079", ''::text)::numeric
                    WHEN NULLIF(rp."RIAD4079", ''::text) IS NOT NULL AND NULLIF(rd."RIAD4079", ''::text) IS NOT NULL THEN NULLIF(ri."RIAD4079", ''::text)::numeric + NULLIF(rd."RIAD4079", ''::text)::numeric - NULLIF(rp."RIAD4079", ''::text)::numeric
                    ELSE NULL::numeric
                END / NULLIF(NULLIF(rc."RCON2200", ''::text)::numeric, 0::numeric) * 100::numeric, 4) AS noninterest_income_pct,
            round(NULLIF(ri."RIAD4093", ''::text)::numeric / NULLIF(NULLIF(ri."RIAD4074", ''::text)::numeric + NULLIF(ri."RIAD4079", ''::text)::numeric, 0::numeric) * 100::numeric, 2) AS efficiency_ratio,
            (COALESCE(NULLIF(rc."RCONB987", ''::text)::numeric, 0::numeric) + COALESCE(NULLIF(rc."RCONB989", ''::text)::numeric, 0::numeric)) * 1000::numeric AS brokered_deposits,
            COALESCE(NULLIF(rc."RCONB993", ''::text)::numeric, 0::numeric) * 1000::numeric AS reciprocal_deposits,
            round((COALESCE(NULLIF(rc."RCONB987", ''::text)::numeric, 0::numeric) + COALESCE(NULLIF(rc."RCONB989", ''::text)::numeric, 0::numeric)) / NULLIF(NULLIF(rc."RCON2200", ''::text)::numeric, 0::numeric) * 100::numeric, 2) AS brokered_deposits_pct
           FROM "raw_schedule_RI" ri
             JOIN "raw_schedule_RC" rc ON rc."IDRSSD" = ri."IDRSSD" AND rc.period = ri.period
             LEFT JOIN "raw_schedule_RI" rp ON rp."IDRSSD" = ri."IDRSSD" AND rp.period = (ri.period - '1 year'::interval)::date
             LEFT JOIN "raw_schedule_RI" rd ON rd."IDRSSD" = ri."IDRSSD" AND rd.period = make_date(EXTRACT(year FROM ri.period)::integer - 1, 12, 31)
          WHERE ri."IDRSSD" IS NOT NULL
        ), bank_ubpr AS (
         SELECT 'bank_'::text || u."ID RSSD" AS inst_key,
            u.period,
            NULLIF(u."UBPRE013", ''::text)::numeric AS roa,
            NULLIF(u."UBPRE018", ''::text)::numeric AS nim,
            NULLIF(u."UBPR7316", ''::text)::numeric AS asset_growth_pct,
            NULLIF(u."UBPRE027", ''::text)::numeric AS loan_growth_pct,
            NULLIF(u."UBPR7414", ''::text)::numeric AS noncurrent_assets_pct,
            NULLIF(u."UBPRE549", ''::text)::numeric AS nonperforming_pct,
            NULLIF(u."UBPRE600", ''::text)::numeric AS loans_to_deposits_pct,
            NULLIF(u."UBPRD486", ''::text)::numeric AS tier1_capital_pct,
            NULLIF(u."UBPR7408", ''::text)::numeric AS tier1_change_pct,
            NULLIF(u."UBPRE005", ''::text)::numeric AS overhead_pct
           FROM "raw_UBPR" u
          WHERE u."ID RSSD" IS NOT NULL AND u."ID RSSD" <> ''::text
        ), bank_with_history AS (
         SELECT b.inst_key,
            b.institution_id,
            b.institution_type,
            b.period,
            b.total_assets,
            b.total_deposits,
            b.total_equity,
            b.net_interest_income,
            b.noninterest_income,
            b.total_interest_expense,
            b.noninterest_expense,
            b.net_income,
            b.save_interest_pct,
            b.cd_interest_pct,
            b.cost_of_funds_pct,
            b.noninterest_income_pct,
            b.efficiency_ratio,
            b.brokered_deposits,
            b.reciprocal_deposits,
            b.brokered_deposits_pct,
            round((b.total_deposits - bq.total_deposits) / NULLIF(bq.total_deposits, 0::numeric) * 100::numeric, 2) AS dep_qoq_pct,
            round((b.total_deposits - by2.total_deposits) / NULLIF(by2.total_deposits, 0::numeric) * 100::numeric, 2) AS dep_yoy_pct,
            round((b.net_income - by2.net_income) / NULLIF(abs(by2.net_income), 0::numeric) * 100::numeric, 2) AS net_income_yoy_pct,
            round(b.cost_of_funds_pct - bq.cost_of_funds_pct, 4) AS cof_qoq_change,
            round(b.cost_of_funds_pct - by2.cost_of_funds_pct, 4) AS cof_yoy_change,
            round(b.save_interest_pct - by2.save_interest_pct, 4) AS save_rate_yoy_change,
            round(b.cd_interest_pct - by2.cd_interest_pct, 4) AS cd_rate_yoy_change,
            round(b.efficiency_ratio - by2.efficiency_ratio, 2) AS efficiency_ratio_yoy_change
           FROM bank_base b
             LEFT JOIN bank_base bq ON bq.inst_key = b.inst_key AND bq.period = (date_trunc('month'::text, b.period::timestamp without time zone) - '2 mons'::interval - '1 day'::interval)::date
             LEFT JOIN LATERAL ( SELECT by_.total_deposits,
                    by_.net_income,
                    by_.cost_of_funds_pct,
                    by_.save_interest_pct,
                    by_.cd_interest_pct,
                    by_.efficiency_ratio
                   FROM bank_base by_
                  WHERE by_.inst_key = b.inst_key AND by_.period >= (b.period - '1 year 3 mons'::interval)::date AND by_.period <= (b.period - '9 mons'::interval)::date
                  ORDER BY (abs(EXTRACT(epoch FROM by_.period::timestamp without time zone - (b.period - '1 year'::interval))))
                 LIMIT 1) by2 ON true
        ), cu_base AS (
         SELECT 'cu_'::text || c."CU_NUMBER" AS inst_key,
            c."CU_NUMBER" AS institution_id,
            'cu'::text AS institution_type,
            to_date(split_part(c."CYCLE_DATE", ' '::text, 1), 'MM/DD/YYYY'::text) AS period,
            NULLIF(c."ACCT_010", ''::text)::numeric AS total_assets,
            NULLIF(c."ACCT_018", ''::text)::numeric AS total_deposits,
            NULLIF(c."ACCT_940", ''::text)::numeric AS total_equity,
            NULLIF(c."ACCT_084", ''::text)::numeric AS net_income,
            NULLIF(c."ACCT_380", ''::text)::numeric AS noninterest_expense,
            COALESCE(NULLIF(c."ACCT_550", ''::text)::numeric, 0::numeric) + COALESCE(NULLIF(c."ACCT_551", ''::text)::numeric, 0::numeric) AS total_interest_expense,
                CASE
                    WHEN NULLIF(c."ACCT_018", ''::text)::numeric > 0::numeric THEN LEAST(round((COALESCE(NULLIF(c."ACCT_550", ''::text)::numeric, 0::numeric) + COALESCE(NULLIF(c."ACCT_551", ''::text)::numeric, 0::numeric)) / NULLIF(c."ACCT_018", ''::text)::numeric * (12.0 / EXTRACT(month FROM to_date(split_part(c."CYCLE_DATE", ' '::text, 1), 'MM/DD/YYYY'::text))) * 100::numeric, 4), 5.0)
                    ELSE NULL::numeric
                END AS cost_of_funds_pct,
                CASE
                    WHEN NULLIF(c."ACCT_018", ''::text)::numeric > 0::numeric THEN LEAST(round(COALESCE(NULLIF(c."ACCT_550", ''::text)::numeric, 0::numeric) / NULLIF(c."ACCT_018", ''::text)::numeric * (12.0 / EXTRACT(month FROM to_date(split_part(c."CYCLE_DATE", ' '::text, 1), 'MM/DD/YYYY'::text))) * 100::numeric, 4), 5.0)
                    ELSE NULL::numeric
                END AS save_interest_pct,
                CASE
                    WHEN NULLIF(c."ACCT_902", ''::text)::numeric > 0::numeric THEN LEAST(round(COALESCE(NULLIF(c."ACCT_551", ''::text)::numeric, 0::numeric) / NULLIF(c."ACCT_902", ''::text)::numeric * (12.0 / EXTRACT(month FROM to_date(split_part(c."CYCLE_DATE", ' '::text, 1), 'MM/DD/YYYY'::text))) * 100::numeric, 4), 8.0)
                    ELSE NULL::numeric
                END AS cd_interest_pct,
                CASE
                    WHEN NULLIF(c."ACCT_010", ''::text)::numeric > 0::numeric THEN round(NULLIF(c."ACCT_084", ''::text)::numeric / NULLIF(c."ACCT_010", ''::text)::numeric * (12.0 / EXTRACT(month FROM to_date(split_part(c."CYCLE_DATE", ' '::text, 1), 'MM/DD/YYYY'::text))) * 100::numeric, 4)
                    ELSE NULL::numeric
                END AS roa,
                CASE
                    WHEN NULLIF(c."ACCT_010", ''::text)::numeric > 0::numeric THEN round(NULLIF(c."ACCT_940", ''::text)::numeric / NULLIF(c."ACCT_010", ''::text)::numeric * 100::numeric, 2)
                    ELSE NULL::numeric
                END AS net_worth_ratio_pct,
                CASE
                    WHEN NULLIF(c."ACCT_018", ''::text)::numeric > 0::numeric THEN round(NULLIF(c."ACCT_013", ''::text)::numeric / NULLIF(c."ACCT_018", ''::text)::numeric * 100::numeric, 2)
                    ELSE NULL::numeric
                END AS loans_to_deposits_pct,
                CASE
                    WHEN (NULLIF(c."ACCT_084", ''::text)::numeric + NULLIF(c."ACCT_380", ''::text)::numeric) > 0::numeric THEN round(NULLIF(c."ACCT_380", ''::text)::numeric / (NULLIF(c."ACCT_084", ''::text)::numeric + NULLIF(c."ACCT_380", ''::text)::numeric) * 100::numeric, 2)
                    ELSE NULL::numeric
                END AS efficiency_ratio
           FROM "Raw_cu_fs220" c
          WHERE c."CU_NUMBER" IS NOT NULL
        ), cu_with_history AS (
         SELECT c.inst_key,
            c.institution_id,
            c.institution_type,
            c.period,
            c.total_assets,
            c.total_deposits,
            c.total_equity,
            c.net_income,
            c.noninterest_expense,
            c.total_interest_expense,
            c.cost_of_funds_pct,
            c.save_interest_pct,
            c.cd_interest_pct,
            c.roa,
            c.net_worth_ratio_pct,
            c.loans_to_deposits_pct,
            c.efficiency_ratio,
            round((c.total_deposits - cq.total_deposits) / NULLIF(cq.total_deposits, 0::numeric) * 100::numeric, 2) AS dep_qoq_pct,
            round((c.total_deposits - cy.total_deposits) / NULLIF(cy.total_deposits, 0::numeric) * 100::numeric, 2) AS dep_yoy_pct,
            round((c.net_income - cy.net_income) / NULLIF(abs(cy.net_income), 0::numeric) * 100::numeric, 2) AS net_income_yoy_pct,
            round(c.cost_of_funds_pct - cq.cost_of_funds_pct, 4) AS cof_qoq_change,
            round(c.cost_of_funds_pct - cy.cost_of_funds_pct, 4) AS cof_yoy_change,
            round(c.save_interest_pct - cy.save_interest_pct, 4) AS save_rate_yoy_change,
            round(c.cd_interest_pct - cy.cd_interest_pct, 4) AS cd_rate_yoy_change,
            round(c.efficiency_ratio - cy.efficiency_ratio, 2) AS efficiency_ratio_yoy_change
           FROM cu_base c
             LEFT JOIN cu_base cq ON cq.inst_key = c.inst_key AND cq.period = (date_trunc('month'::text, c.period::timestamp without time zone) - '2 mons'::interval - '1 day'::interval)::date
             LEFT JOIN LATERAL ( SELECT cy_.total_deposits,
                    cy_.net_income,
                    cy_.cost_of_funds_pct,
                    cy_.save_interest_pct,
                    cy_.cd_interest_pct,
                    cy_.efficiency_ratio
                   FROM cu_base cy_
                  WHERE cy_.inst_key = c.inst_key AND cy_.period >= (c.period - '1 year 3 mons'::interval)::date AND cy_.period <= (c.period - '9 mons'::interval)::date
                  ORDER BY (abs(EXTRACT(epoch FROM cy_.period::timestamp without time zone - (c.period - '1 year'::interval))))
                 LIMIT 1) cy ON true
        )
 SELECT b.inst_key,
    b.institution_id,
    b.institution_type,
    b.period,
    b.total_assets,
    b.total_deposits,
    b.total_equity,
    b.net_interest_income,
    b.noninterest_income,
    b.total_interest_expense,
    b.noninterest_expense,
    b.net_income,
    b.save_interest_pct,
    b.cd_interest_pct,
    b.cost_of_funds_pct,
    b.noninterest_income_pct,
    b.efficiency_ratio,
    u.roa,
    u.nim,
    u.asset_growth_pct,
    u.loan_growth_pct,
    u.noncurrent_assets_pct,
    u.nonperforming_pct,
    u.loans_to_deposits_pct,
    u.tier1_capital_pct,
    u.tier1_change_pct,
    u.overhead_pct,
    NULL::numeric AS net_worth_ratio_pct,
    b.dep_qoq_pct,
    b.dep_yoy_pct,
    b.net_income_yoy_pct,
    b.cof_qoq_change,
    b.cof_yoy_change,
    b.save_rate_yoy_change,
    b.cd_rate_yoy_change,
    b.efficiency_ratio_yoy_change,
    b.brokered_deposits,
    b.reciprocal_deposits,
    b.brokered_deposits_pct
   FROM bank_with_history b
     LEFT JOIN bank_ubpr u ON u.inst_key = b.inst_key AND u.period = (( SELECT max(u2.period) AS max
           FROM bank_ubpr u2
          WHERE u2.inst_key = b.inst_key AND u2.period <= b.period))
UNION ALL
 SELECT c.inst_key,
    c.institution_id,
    c.institution_type,
    c.period,
    c.total_assets,
    c.total_deposits,
    c.total_equity,
    NULL::numeric AS net_interest_income,
    NULL::numeric AS noninterest_income,
    c.total_interest_expense,
    c.noninterest_expense,
    c.net_income,
    c.save_interest_pct,
    c.cd_interest_pct,
    c.cost_of_funds_pct,
    NULL::numeric AS noninterest_income_pct,
    c.efficiency_ratio,
    c.roa,
    NULL::numeric AS nim,
    NULL::numeric AS asset_growth_pct,
    NULL::numeric AS loan_growth_pct,
    NULL::numeric AS noncurrent_assets_pct,
    NULL::numeric AS nonperforming_pct,
    c.loans_to_deposits_pct,
    NULL::numeric AS tier1_capital_pct,
    NULL::numeric AS tier1_change_pct,
    NULL::numeric AS overhead_pct,
    c.net_worth_ratio_pct,
    c.dep_qoq_pct,
    c.dep_yoy_pct,
    c.net_income_yoy_pct,
    c.cof_qoq_change,
    c.cof_yoy_change,
    c.save_rate_yoy_change,
    c.cd_rate_yoy_change,
    c.efficiency_ratio_yoy_change,
    NULL::numeric AS brokered_deposits,
    NULL::numeric AS reciprocal_deposits,
    NULL::numeric AS brokered_deposits_pct
   FROM cu_with_history c;
