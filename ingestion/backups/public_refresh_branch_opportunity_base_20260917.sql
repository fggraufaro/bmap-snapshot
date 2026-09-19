-- Backup of the STALE, BUGGY public.refresh_branch_opportunity_base() before it was
-- dropped as part of a56 (2026-09-17).
--
-- Bug (a56): this version's tmp_branch_yoy is built purely from raw.raw_sod (FDIC
-- Summary of Deposits, banks only). raw_sod has zero credit-union rows, so the
-- later `FROM geo.branches_master_v2 bm JOIN tmp_branch_yoy by_ ON by_.uninumbr =
-- bm.branch_id::text` is an inner join that silently drops all 20,351 CU branches
-- from analytics.branch_opportunity_base if this version ever runs. Also carries
-- the older fixed-10mi radius (geo.branch_competitors_10mi_v2 /
-- geo.branch_radius_stats_10mi_v2) and the pre-16-play-matrix ad-hoc campaign
-- labels, both since superseded in analytics.refresh_branch_opportunity_base().
--
-- analytics.refresh_branch_opportunity_base() is the correct, current version:
-- it already unions in CU branches (deposits from geo.branches_master_v2.
-- depsum_branch, YoY from analytics.bank_financial_snapshot_latest.dep_yoy_pct
-- applied per institution), uses the tiered adaptive radius, and the authoritative
-- 16-play matrix. That is the one all callers (including ingestion/run_all.py)
-- should target via Content-Profile: analytics.
--
-- This file exists only as a historical record of the dead logic being removed,
-- per the project's backup-before-destructive-op rule. Do not restore this
-- without first re-fixing the CU/radius/campaign issues it contains.

CREATE OR REPLACE FUNCTION public.refresh_branch_opportunity_base()
 RETURNS void
 LANGUAGE plpgsql
AS $function$
BEGIN

    RAISE NOTICE 'Step 1/10: Building income YoY...';
    CREATE TEMP TABLE tmp_income ON COMMIT DROP AS
    WITH cleaned AS (
        SELECT SPLIT_PART("Geographic Area Name", ' ', 2) AS zip,
            NULLIF(REGEXP_REPLACE("income", '[^0-9]', '', 'g'), '') AS income_clean,
            "YEAR"::integer AS year
        FROM raw.raw_income WHERE "income" NOT IN ('-', '') AND "income" IS NOT NULL
    ),
    latest AS (SELECT zip, income_clean::numeric AS household_income FROM cleaned WHERE year = (SELECT MAX(year) FROM cleaned)),
    prev   AS (SELECT zip, income_clean::numeric AS household_income_prev FROM cleaned WHERE year = (SELECT MAX(year) - 1 FROM cleaned))
    SELECT l.zip, l.household_income,
        CASE WHEN p.household_income_prev IS NULL OR p.household_income_prev = 0 THEN NULL
             ELSE ROUND((l.household_income - p.household_income_prev) / p.household_income_prev, 6)
        END AS yoy_income_growth
    FROM latest l LEFT JOIN prev p ON p.zip = l.zip;

    RAISE NOTICE 'Step 2/10: Building population YoY...';
    CREATE TEMP TABLE tmp_population ON COMMIT DROP AS
    WITH latest AS (
        SELECT SPLIT_PART("Geographic Area Name", ' ', 2) AS zip, total AS total_population
        FROM raw.raw_population WHERE "YEAR" = (SELECT MAX("YEAR") FROM raw.raw_population)
    ),
    prev AS (
        SELECT SPLIT_PART("Geographic Area Name", ' ', 2) AS zip, total AS total_population_prev
        FROM raw.raw_population WHERE "YEAR" = (SELECT (MAX("YEAR")::integer - 1)::text FROM raw.raw_population)
    )
    SELECT l.zip, l.total_population,
        CASE WHEN p.total_population_prev IS NULL OR p.total_population_prev = 0 THEN NULL
             ELSE ROUND((l.total_population - p.total_population_prev)::numeric / p.total_population_prev, 6)
        END AS yoy_pop_growth
    FROM latest l LEFT JOIN prev p ON p.zip = l.zip;

    RAISE NOTICE 'Step 3/10: Building branch YoY deposits from raw_sod...';
    CREATE TEMP TABLE tmp_branch_yoy ON COMMIT DROP AS
    WITH latest AS (
        SELECT DISTINCT ON ("UNINUMBR") "UNINUMBR" AS uninumbr, "DEPSUMBR"::numeric * 1000 AS latest_dep
        FROM raw.raw_sod
        WHERE "YEAR" = (SELECT MAX("YEAR") FROM raw.raw_sod)
          AND "STALPBR" NOT IN ('AK','AS','GU','HI','MP','PR','VI')
          AND "DEPSUMBR" IS NOT NULL AND "DEPSUMBR"::numeric > 0
        ORDER BY "UNINUMBR"
    ),
    prev AS (
        SELECT DISTINCT ON ("UNINUMBR") "UNINUMBR" AS uninumbr, "DEPSUMBR"::numeric * 1000 AS prev_dep
        FROM raw.raw_sod WHERE "YEAR" = (SELECT (MAX("YEAR")::integer - 1)::text FROM raw.raw_sod)
        ORDER BY "UNINUMBR"
    )
    SELECT l.uninumbr, l.latest_dep,
        GREATEST(-1.0, LEAST(1.0,
            CASE WHEN p.prev_dep IS NULL OR p.prev_dep = 0 THEN NULL
                 ELSE ROUND((l.latest_dep - p.prev_dep)::numeric / p.prev_dep, 6)
            END)) AS yoy_deposits
    FROM latest l LEFT JOIN prev p ON p.uninumbr = l.uninumbr;

    RAISE NOTICE 'Step 4/10: Building avg competitor YoY...';
    CREATE TEMP TABLE tmp_avg_comp_yoy ON COMMIT DROP AS
    SELECT c.my_branch_id AS branch_id,
        GREATEST(-1.0, LEAST(1.0, ROUND(AVG(y.yoy_deposits), 6))) AS avg_comp_yoy
    FROM geo.branch_competitors_10mi_v2 c
    JOIN tmp_branch_yoy y ON y.uninumbr = c.competitor_branch_uninumbr::text
    WHERE y.yoy_deposits IS NOT NULL GROUP BY c.my_branch_id;

    RAISE NOTICE 'Step 5/10: Building ZHVI per ZIP...';
    CREATE TEMP TABLE tmp_zhvi ON COMMIT DROP AS
    WITH latest AS (SELECT DISTINCT ON (zip) zip, zhvi, period FROM raw.raw_zhvi ORDER BY zip, period DESC),
    prev AS (
        SELECT DISTINCT ON (zip) zip, zhvi AS zhvi_prev FROM raw.raw_zhvi
        WHERE period < (SELECT MAX(period) - INTERVAL '12 months' FROM raw.raw_zhvi)
        ORDER BY zip, period DESC
    )
    SELECT l.zip, l.zhvi, l.period AS zhvi_period,
        CASE WHEN p.zhvi_prev IS NULL OR p.zhvi_prev = 0 THEN NULL
             ELSE ROUND((l.zhvi - p.zhvi_prev) / p.zhvi_prev * 100, 2)
        END AS zhvi_yoy_pct
    FROM latest l LEFT JOIN prev p ON p.zip = l.zip;

    RAISE NOTICE 'Step 5b/10: Pre-aggregating radius stats...';
    CREATE TEMP TABLE tmp_radius_stats ON COMMIT DROP AS
    SELECT DISTINCT ON (my_branch_id) my_branch_id, branches_in_radius,
        bank_branches_in_radius, cu_branches_in_radius,
        total_deposits_10mi::numeric, cu_deposit_share_pct::numeric,
        avg_deposits_per_branch_10mi::numeric
    FROM geo.branch_radius_stats_10mi_v2 ORDER BY my_branch_id;

    RAISE NOTICE 'Step 6/10: Rebuilding branch_opportunity_base...';
    TRUNCATE analytics.branch_opportunity_base;

    INSERT INTO analytics.branch_opportunity_base (
        uninumbr, rssdid, namefull, namebr, addresbr, citybr,
        stalpbr, zipbr, sims_latitude, sims_longitude, insured, year,
        uni_branch, latest_dep, yoy_deposits, household_income,
        yoy_income_growth, total_population, yoy_pop_growth,
        market_growth_score, branch_density_10mi, branches_in_radius,
        bank_branches_in_radius, cu_branches_in_radius,
        total_deposits_10mi, cu_deposit_share_pct, inverted_density,
        avg_comp_yoy, rel_branch_growth, zhvi, zhvi_yoy_pct, zhvi_period,
        smb_establishments, estab_per_1k_pop, smb_index, smb_zone,
        institution_type, inst_key
    )
    SELECT bm.branch_id, bm.bank_id, bm.bank_name, bm.branch_name,
        bm.address, bm.city, bm.state, bm.zip, bm.lat, bm.lon,
        bm.insured, bm.year,
        bm.branch_id::text || '--' || bm.branch_name,
        by_.latest_dep, by_.yoy_deposits,
        inc.household_income, inc.yoy_income_growth,
        pop.total_population, pop.yoy_pop_growth,

        -- ── UPDATED: market_growth_score now 40% income + 35% population + 25% ZHVI ──
        -- Previous: 0.5 * income_yoy + 0.5 * population_yoy
        -- ZHVI YoY (%) converted to decimal (÷100) to match income/population scale
        CASE WHEN inc.yoy_income_growth IS NULL
              AND pop.yoy_pop_growth IS NULL
              AND zh.zhvi_yoy_pct IS NULL THEN NULL
             ELSE ROUND(
                 0.40 * COALESCE(inc.yoy_income_growth, 0)
               + 0.35 * COALESCE(pop.yoy_pop_growth, 0)
               + 0.25 * COALESCE(zh.zhvi_yoy_pct / 100.0, 0)
             , 6) END,

        rs.avg_deposits_per_branch_10mi, rs.branches_in_radius,
        rs.bank_branches_in_radius, rs.cu_branches_in_radius,
        rs.total_deposits_10mi, rs.cu_deposit_share_pct,
        CASE WHEN rs.avg_deposits_per_branch_10mi IS NULL OR rs.avg_deposits_per_branch_10mi = 0 THEN NULL
             ELSE ROUND(1.0 / rs.avg_deposits_per_branch_10mi, 10) END,
        ac.avg_comp_yoy,
        CASE WHEN by_.yoy_deposits IS NULL OR ac.avg_comp_yoy IS NULL THEN NULL
             ELSE ROUND(by_.yoy_deposits - ac.avg_comp_yoy, 6) END,
        zh.zhvi, zh.zhvi_yoy_pct, zh.zhvi_period,
        smb.smb_establishments, smb.estab_per_1k_pop, smb.smb_index, smb.smb_zone,
        bm.institution_type,
        CASE WHEN bm.institution_type = 'bank' THEN 'bank_' || bm.bank_id::text
             WHEN bm.institution_type = 'cu'   THEN 'cu_'   || bm.bank_id::text
             ELSE bm.bank_id::text END
    FROM geo.branches_master_v2 bm
    JOIN  tmp_branch_yoy     by_ ON by_.uninumbr    = bm.branch_id::text
    LEFT JOIN tmp_income     inc ON inc.zip          = bm.zip
    LEFT JOIN tmp_population pop ON pop.zip          = bm.zip
    LEFT JOIN tmp_radius_stats rs ON rs.my_branch_id = bm.branch_id
    LEFT JOIN tmp_avg_comp_yoy ac ON ac.branch_id    = bm.branch_id
    LEFT JOIN tmp_zhvi         zh ON zh.zip          = bm.zip
    LEFT JOIN public.vw_smb_index_by_zip smb ON smb.zip = bm.zip;

    RAISE NOTICE 'Step 7/10: Normalising scores with deposit size weight...';

    UPDATE analytics.branch_opportunity_base
    SET yoy_deposits = GREATEST(-1.0, LEAST(1.0, yoy_deposits))
    WHERE yoy_deposits IS NOT NULL;

    -- Deposit size norm (log scale, state-relative)
    WITH state_dep AS (
        SELECT stalpbr, MIN(LN(NULLIF(latest_dep,0))) AS min_ln, MAX(LN(NULLIF(latest_dep,0))) AS max_ln
        FROM analytics.branch_opportunity_base WHERE latest_dep > 0 GROUP BY stalpbr
    )
    UPDATE analytics.branch_opportunity_base b
    SET deposit_size_norm = CASE WHEN s.max_ln = s.min_ln THEN 50
        ELSE ROUND(((LN(b.latest_dep) - s.min_ln) / (s.max_ln - s.min_ln) * 100)::numeric, 2) END
    FROM state_dep s WHERE b.stalpbr = s.stalpbr AND b.latest_dep > 0;

    WITH winsor AS (
        SELECT PERCENTILE_CONT(0.05) WITHIN GROUP (ORDER BY inverted_density)  AS inv_p05,
               PERCENTILE_CONT(0.95) WITHIN GROUP (ORDER BY inverted_density)  AS inv_p95,
               PERCENTILE_CONT(0.05) WITHIN GROUP (ORDER BY rel_branch_growth) AS rel_p05,
               PERCENTILE_CONT(0.95) WITHIN GROUP (ORDER BY rel_branch_growth) AS rel_p95
        FROM analytics.branch_opportunity_base WHERE inverted_density IS NOT NULL AND rel_branch_growth IS NOT NULL
    ),
    state_mm AS (
        SELECT stalpbr, MIN(market_growth_score) AS min_mgs, MAX(market_growth_score) AS max_mgs
        FROM analytics.branch_opportunity_base WHERE market_growth_score IS NOT NULL GROUP BY stalpbr
    ),
    computed AS (
        SELECT b.uninumbr,
            CASE WHEN sm.max_mgs = sm.min_mgs OR b.market_growth_score IS NULL THEN 0::numeric
                 ELSE ROUND(((b.market_growth_score - sm.min_mgs)/(sm.max_mgs - sm.min_mgs)*100)::numeric,2) END AS mgn,
            CASE WHEN b.inverted_density IS NULL THEN NULL::numeric
                 WHEN w.inv_p95 = w.inv_p05 THEN 0::numeric
                 ELSE ROUND(((LEAST(w.inv_p95,GREATEST(w.inv_p05,b.inverted_density::float))-w.inv_p05)/(w.inv_p95-w.inv_p05)*100)::numeric,2) END AS idn,
            CASE WHEN b.rel_branch_growth IS NULL THEN NULL::numeric
                 WHEN w.rel_p95 = w.rel_p05 THEN 0::numeric
                 ELSE ROUND(((LEAST(w.rel_p95,GREATEST(w.rel_p05,b.rel_branch_growth::float))-w.rel_p05)/(w.rel_p95-w.rel_p05)*100)::numeric,2) END AS rgn
        FROM analytics.branch_opportunity_base b CROSS JOIN winsor w LEFT JOIN state_mm sm ON sm.stalpbr = b.stalpbr
    )
    UPDATE analytics.branch_opportunity_base b
    SET market_growth_normalized = c.mgn,
        inv_density_norm_winsor  = c.idn,
        rel_growth_norm          = c.rgn,
        opportunity_score = ROUND(
            0.25 * COALESCE(c.mgn,0)
          + 0.30 * COALESCE(c.rgn,0)
          + 0.25 * COALESCE(c.idn,0)
          + 0.20 * COALESCE(b.deposit_size_norm,0), 2)
    FROM computed c WHERE b.uninumbr = c.uninumbr;

    RAISE NOTICE 'Step 8/10: Zones from community bank thresholds, applied to all...';
    WITH community_scores AS (
        SELECT b.stalpbr, b.opportunity_score
        FROM analytics.branch_opportunity_base b
        LEFT JOIN analytics.bank_financial_snapshot_latest f ON f.inst_key = b.inst_key
        WHERE COALESCE(f.total_assets, 0) < 10000000000
        AND b.opportunity_score IS NOT NULL
    ),
    state_quartiles AS (
        SELECT stalpbr,
            PERCENTILE_CONT(0.25) WITHIN GROUP (ORDER BY opportunity_score) AS p25,
            PERCENTILE_CONT(0.50) WITHIN GROUP (ORDER BY opportunity_score) AS p50,
            PERCENTILE_CONT(0.75) WITHIN GROUP (ORDER BY opportunity_score) AS p75
        FROM community_scores GROUP BY stalpbr
    )
    UPDATE analytics.branch_opportunity_base b
    SET opportunity_zone = CASE
        WHEN b.opportunity_score >= q.p75 THEN 'Invest'
        WHEN b.opportunity_score >= q.p50 THEN 'Analyze'
        WHEN b.opportunity_score >= q.p25 THEN 'Defend'
        ELSE 'Justify' END
    FROM state_quartiles q WHERE b.stalpbr = q.stalpbr AND b.opportunity_score IS NOT NULL;

    RAISE NOTICE 'Step 8b/10: Rebuilding matrix quadrants...';
    WITH state_zhvi AS (
        SELECT stalpbr, MIN(zhvi_yoy_pct) AS min_z, MAX(zhvi_yoy_pct) AS max_z
        FROM analytics.branch_opportunity_base WHERE zhvi_yoy_pct IS NOT NULL GROUP BY stalpbr
    )
    UPDATE analytics.branch_opportunity_base b
    SET zhvi_yoy_norm = CASE WHEN s.max_z = s.min_z THEN 50
        ELSE ROUND(((b.zhvi_yoy_pct-s.min_z)/(s.max_z-s.min_z)*100)::numeric,2) END
    FROM state_zhvi s WHERE b.stalpbr = s.stalpbr AND b.zhvi_yoy_pct IS NOT NULL;

    UPDATE analytics.branch_opportunity_base
    SET market_attractiveness = ROUND(0.60*COALESCE(market_growth_normalized,0)+0.40*COALESCE(zhvi_yoy_norm,50),2),
        branch_performance    = ROUND(0.60*COALESCE(rel_growth_norm,0)+0.40*COALESCE(inv_density_norm_winsor,0),2);

    WITH state_medians AS (
        SELECT stalpbr,
            PERCENTILE_CONT(0.50) WITHIN GROUP (ORDER BY market_attractiveness) AS med_ma,
            PERCENTILE_CONT(0.50) WITHIN GROUP (ORDER BY branch_performance)    AS med_bp
        FROM analytics.branch_opportunity_base
        WHERE market_attractiveness IS NOT NULL AND branch_performance IS NOT NULL GROUP BY stalpbr
    )
    UPDATE analytics.branch_opportunity_base b
    SET matrix_quadrant = CASE
            WHEN b.market_attractiveness >= m.med_ma AND b.branch_performance >= m.med_bp THEN 'Q2 - Invest and Protect'
            WHEN b.market_attractiveness >= m.med_ma AND b.branch_performance <  m.med_bp THEN 'Q1 - Grow and Perform'
            WHEN b.market_attractiveness <  m.med_ma AND b.branch_performance >= m.med_bp THEN 'Q3 - Maintain and Improve'
            ELSE 'Q4 - Rationalize and Exit' END,
        quadrant_number = CASE
            WHEN b.market_attractiveness >= m.med_ma AND b.branch_performance >= m.med_bp THEN 2
            WHEN b.market_attractiveness >= m.med_ma AND b.branch_performance <  m.med_bp THEN 1
            WHEN b.market_attractiveness <  m.med_ma AND b.branch_performance >= m.med_bp THEN 3
            ELSE 4 END
    FROM state_medians m
    WHERE b.stalpbr = m.stalpbr AND b.market_attractiveness IS NOT NULL AND b.branch_performance IS NOT NULL;

    RAISE NOTICE 'Step 9/10: Assigning campaign plays...';
    UPDATE analytics.branch_opportunity_base
    SET campaign = CASE
        WHEN opportunity_zone = 'Invest'   AND quadrant_number = 2 THEN 'Aggressive Acquisition'
        WHEN opportunity_zone = 'Invest'   AND quadrant_number = 1 THEN 'Urgent Competitive Push'
        WHEN opportunity_zone = 'Invest'   AND quadrant_number = 3 THEN 'Capitalize'
        WHEN opportunity_zone = 'Invest'   AND quadrant_number = 4 THEN 'Turnaround'
        WHEN opportunity_zone = 'Analyze'  AND quadrant_number = 2 THEN 'Grow Share'
        WHEN opportunity_zone = 'Analyze'  AND quadrant_number = 1 THEN 'Competitive Defense'
        WHEN opportunity_zone = 'Analyze'  AND quadrant_number = 3 THEN 'Maintain'
        WHEN opportunity_zone = 'Analyze'  AND quadrant_number = 4 THEN 'Stabilize'
        WHEN opportunity_zone = 'Defend'   AND quadrant_number = 3 THEN 'Loyalty/Retention'
        WHEN opportunity_zone = 'Defend'   AND quadrant_number = 2 THEN 'Hold and Watch'
        WHEN opportunity_zone = 'Defend'   AND quadrant_number IN (1,4) THEN 'Efficiency Review'
        WHEN opportunity_zone = 'Justify'  AND quadrant_number IN (3,2) THEN 'No Action Needed'
        WHEN opportunity_zone = 'Justify'  AND quadrant_number IN (1,4) THEN 'Rationalize'
        ELSE 'Monitor'
    END
    WHERE opportunity_zone IS NOT NULL AND quadrant_number IS NOT NULL;

    RAISE NOTICE 'Step 10/10: Rebuilding minmax + SMB...';
    TRUNCATE analytics.branch_opportunity_minmax;
    INSERT INTO analytics.branch_opportunity_minmax
    SELECT stalpbr, MIN(market_growth_score), MAX(market_growth_score),
           MIN(inverted_density), MAX(inverted_density), MIN(rel_branch_growth), MAX(rel_branch_growth)
    FROM analytics.branch_opportunity_base GROUP BY stalpbr;

    UPDATE analytics.branch_opportunity_base b
    SET smb_establishments = s.smb_establishments, estab_per_1k_pop = s.estab_per_1k_pop,
        smb_index = s.smb_index, smb_zone = s.smb_zone
    FROM public.vw_smb_index_by_zip s WHERE b.zipbr = s.zip;

    RAISE NOTICE 'Opportunity base rebuild complete. market_growth_score now: 40pct income + 35pct population + 25pct ZHVI YoY';
END;
$function$
