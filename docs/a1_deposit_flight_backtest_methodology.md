# a1 — Deposit-Flight Backtest: Phase 1 Methodology

Status: **Phase 1 (methodology only) — no backtest has been run under this document.**
Author: Session 3 (Product & Research). Reviewed by: Session 4 (Roadmap Manager) before Phase 2 starts.

This document covers **two separate, independently-evaluated tests**, not one combined score:

- **Test 1** — branch-level competitive density → branch deposit growth (SOD only).
- **Test 2** — institution-level financial health → institution deposit growth (Call Report
  only, banks only).

They are kept separate deliberately: they test different mechanisms at different units of
analysis (branch vs. institution), using different source data on different cadences (annual SOD
vs. quarterly Call Report). Combining them into one composite predictor now, before either has
been validated alone, would make a pass or fail impossible to attribute to either mechanism and
would add an unvalidated aggregation scheme on top of two unproven claims. If both tests pass
independently, whether a combined score outperforms either alone is a legitimate follow-up
question — but it comes after, not instead of, proving the pieces.

Revision note: originally scoped to Test 1 only, on a single transition (2024→2025), because
2023 SOD was only a ~31%-coverage partial subset at the time of writing. On 2026-09-13: (1) a
full 2023 SOD snapshot was loaded (77,770 rows via FDIC's public API, verified independently by
Session 4), which added a second Test 1 transition (2023→2024) per Session 4's direction; (2)
`analytics.bank_financial_snapshot` (a real quarterly panel, distinct from the sparse
`_latest` view) was found to have genuine historical depth for banks, which added Test 2 in this
revision. Neither Test 1's second transition nor any part of Test 2 has been run in any form —
both are still fully prospective.

## Disclosure — read this before anything else

An informal version of Test 1's 2024→2025 transition was already run in this session on
2026-09-11, before this Phase 1/2/3 process existed, and reported in chat to Session 4. Session
4 subsequently found no commit, results artifact, or Found-log entry backing that "done" claim
and reopened the item. This document is written with full knowledge of that earlier informal
run's results for 2024→2025 — I am not pretending to approach that pair blind. Test 1's
2023→2024 pair and all of Test 2 have never been run in any form, informal or otherwise — those
are genuinely prospective. To keep this honest rather than a post-hoc rationalization of numbers
I already like:

- Test 1's 2024→2025 methodology (data periods, predictor definition, join logic) matches what
  was actually run on 2026-09-11, because that methodology was sound on its own merits, not
  because it produced a favorable number — the reasoning for each choice is given independently
  below, and the identical logic is applied to 2023→2024 without modification.
- Both success bars (Test 1 Section 1.3, Test 2 Section 2.3) are set from first principles (what
  would make each claim practically useful to Verlocity), not reverse-engineered from any result.
- Phase 2 will run both tests against whatever the relevant tables contain at that time and
  report the results exactly as they come out, including if any of it now fails the bars below.

---

# Test 1 — Branch-level competitive density → branch deposit growth

## 1.1 The testable claim

**Claim under test:** a branch-level competitive-density signal, computed from data available
at the *start* of a period, predicts that branch's deposit trajectory over the *following*
period — specifically, branches facing higher local competitive density see systematically
worse subsequent deposit growth than branches facing lower density, after controlling for the
branch's own market share (since raw density confounds with urbanicity, which has its own
baseline growth effects unrelated to competition — this confound must be controlled for, not
ignored, or the test is invalid either way).

This is the same underlying thesis `analytics.branch_opportunity_base`'s adaptive-radius
inverse-density term is built on. The backtest exists to check whether that thesis holds up
against real historical outcomes, before treating "predictive model" as a sellable feature.

**Why two transitions, not one:** a single-period pass could be coincidence, or a fit to that
specific period's dynamics (a rate environment, a regional M&A wave, anything idiosyncratic to
that year) rather than a genuine, repeatable competitive-density effect. Testing the same
methodology against two independent period-pairs (2023→2024 and 2024→2025) is what actually
distinguishes a real signal from a one-off. This is also the reason a *combined/pooled*
threshold across both transitions is explicitly rejected as a substitute (see 1.3) — a pooled
test would just be a bigger single sample, not a replication check.

**What would falsify the claim:** if the predictor shows no relationship (or the wrong sign)
with next-period deposit growth once market-share tier is controlled for — in either
transition — the thesis does not hold up under a real historical test and should not be
marketed as predictive without further work.

## 1.2 Exact data

- **Source table:** `raw.raw_sod` (FDIC Summary of Deposits, branch-level, annual).
- **Snapshot as of 2026-09-13, post-2023-reload:**
  | Year | Rows | Distinct branches (`UNINUMBR`) |
  |---|---|---|
  | 2023 | 77,770 | 77,770 |
  | 2024 | 76,742 | 76,742 |
  | 2025 | 76,148 | 76,148 |

  All three years are now full-population snapshots of comparable scale (2023 was previously a
  ~31%-coverage partial subset — 24,020 rows — which is why the original version of this
  document restricted the test to 2024→2025 only; that restriction no longer applies).

- **Two transitions, tested independently and identically:**

  | | Transition A | Transition B |
  |---|---|---|
  | Prediction period | 2023 | 2024 |
  | Outcome period | 2024 | 2025 |

- **Join key:** `UNINUMBR` (branch-level identifier), present in both years of a given
  transition for the same physical branch.
- **Predictor, computed entirely from the prediction-period year (no look-ahead), identically
  for both transitions:**
  - `own_share` = branch's `DEPSUMBR` / total `DEPSUMBR` across all branches in the same county
    (`STCNTYBR`) in the prediction-period year.
  - `competitor_count` = count of other branches (any institution) in the same county in the
    prediction-period year.
- **Outcome, computed from the outcome-period year:**
  - `yoy_growth` = (`DEPSUMBR`<sub>outcome</sub> − `DEPSUMBR`<sub>prediction</sub>) /
    `DEPSUMBR`<sub>prediction</sub>.
- **Exclusions, decided here and applied identically to both transitions regardless of result:**
  - `DEPSUMBR`<sub>prediction</sub> `< $1,000` (in thousands, i.e. <$1M actual) excluded — tiny
    denominators produce meaningless percentage swings.
  - `DEPSUMBR`<sub>outcome</sub> `< 5% of DEPSUMBR`<sub>prediction</sub> excluded — this is a
    branch-code retirement/merger artifact (deposits going to effectively zero), not organic
    deposit flight, and must not be counted as a "prediction success" for either side of the
    test.
  - `yoy_growth` winsorized to [−90%, +200%] before any averaging, to prevent a small number of
    extreme values from dominating a mean (medians are reported alongside means regardless).
- **Stratification, identical for both transitions:** branches split into terciles by
  `own_share` and, within each share tercile, terciles by `competitor_count` — 9 cells per
  transition, 18 cells total. This is necessary because competitor count correlates with
  urbanicity, which has its own baseline growth rate unrelated to competition; a population-wide
  correlation without this control is expected to wash out real effects in different directions
  across strata (this was in fact observed on 9/11 for transition B: population correlation ≈ 0
  on the raw pooled data).

## 1.3 Success bar — decided now, before Phase 2 runs

**Both transitions must independently satisfy all three criteria below.** A pass on only one
transition does **not** support the claim — it does not rule out the coincidence/period-specific
risk that is the entire reason a second transition was added (1.1). There is no combined or
averaged threshold across the two transitions; each is graded on its own, and both must pass.

For a given transition, the claim is considered **supported** on that transition only if *all*
of the following hold:

1. **Directional monotonicity:** within at least 2 of the 3 share terciles, median `yoy_growth`
   decreases monotonically as `competitor_count` tercile increases (low → mid → high
   competition). A single non-monotonic cell in a tercile with fewer than 500 branches does not
   break this (small-sample noise is expected there).
2. **Practical magnitude:** the spread between the low-competition and high-competition median
   growth, within at least one share tercile, is **≥ 3 percentage points**. Below that, the
   effect is statistically present but too small to be a sellable predictive signal.
3. **Not an artifact of closures:** the result must hold (same direction, magnitude within 1
   point) after excluding branches where the outcome-period deposits fell below 5% of the
   prediction-period deposits (1.2's exclusion) — i.e., the effect is not being driven by branch
   closures/mergers miscounted as organic flight.

**Overall Test 1 outcome:**
- **Supported** — both transitions pass all three criteria independently.
- **Not supported** — either transition fails any criterion, including if the population-wide
  (unstratified) correlation is the only place an effect shows up in that transition, since that
  would indicate the "signal" is actually a confound (urbanicity, market size) rather than a
  genuine competitive-density effect.
- **Mixed** (one transition passes, one fails) is treated as **not supported** for the purpose of
  the BlastPoint-facing "predictive" pitch and a2's evidence section — see 1.1 for why a single
  passing transition isn't sufficient on its own. A mixed result is still worth reporting in full
  (which transition passed, which didn't, and by how much) since it's informative for follow-up
  work, just not sufficient to claim the predictive thesis is validated.

---

# Test 2 — Institution-level financial health → institution deposit growth

Test 2 runs **four parallel predictor variants** — ROA, loans-to-deposits ratio (LDR), brokered
deposits %, and NIM — sharing identical data, transitions, exclusions, outcome definition, and
size-tercile control (2.2), and identical success-bar criteria (2.3). Only the specific "health"
field being tercile-stratified against size changes between the four. They are reported and
evaluated independently, not combined into one Test 2 verdict — see 2.1 for why, and 2.3 for the
multiple-comparisons consequence of testing four instead of one.

## 2.1 The testable claims

**Shared claim under test:** an institution's own financial health, measured at the *start* of a
period by one of four fields (ROA, LDR, brokered deposits %, NIM), predicts that institution's
deposit growth over the *following* period, after controlling for institution size (since larger
and smaller institutions may have systematically different baseline growth unrelated to
fundamentals — the same confound logic as Test 1's market-share control, applied here to size
instead). One shared hypothesis direction across all four, consistent with how
`branch_target_competitors`' `vuln_score` already treats these signals: **weaker fundamentals
predict worse subsequent deposit growth.** Concretely — lower ROA, higher LDR, higher brokered
deposits %, or lower NIM at the start of a period each predict worse deposit growth over the
following period, controlling for size.

**Why these four, and why one shared direction instead of four separate hypotheses:**
- **ROA** — the most direct match to what `vuln_score` already uses; a general profitability
  summary. Weakest mechanistic link to deposits specifically of the four, but the natural
  baseline given it's already load-bearing elsewhere in the codebase.
- **Loans-to-deposits ratio** — more directly deposit-relevant than ROA: a bank running high LDR
  has less deposit cushion relative to its loan book and is more dependent on continuing to
  attract/retain deposits to fund growth.
- **Brokered deposits %** — the sharpest of the four. High reliance on brokered ("hot money")
  deposits is a well-known real stress marker (this is the same category of signal that flagged
  SVB/Signature-style funding fragility before those failures) — a bank leaning on brokered
  deposits is often plugging a gap because its organic deposit growth is already weak, which
  makes it close to a direct leading indicator rather than a proxy.
- **NIM** — margin compression can reflect a bank paying up for deposits to retain them; adjacent
  to ROA but more specific to funding-cost pressure.

This is a different mechanism from Test 1 across all four: not "is there more competition
nearby," but "is this specific institution financially strong enough to hold and grow its own
deposits." Using the institution's own trajectory is the cleanest available test case for this
family of signals (predicting a *competitor's* future deposits from the outside, which is what
`vuln_score` actually has to do in production, is a harder, noisier version of the same
underlying claim; this test validates the claim in its easiest, most direct form first).

**Why banks only, not credit unions:** `analytics.bank_financial_snapshot` has only one quarter
of data for credit unions (2025-09-30), with no second point to measure a transition against —
checked directly before writing this. CU-side testing of this claim is a separate follow-up item,
not something this document can do yet.

**Why two transitions, not one:** same reasoning as Test 1 (1.1) — a single-quarter pass could be
an artifact of that specific quarter's rate environment or a one-off sector event, not a
repeatable signal. As with Test 1, there is no combined/pooled threshold across the two
transitions for any of the four predictors — each predictor is graded independently per
transition, and both transitions must pass for that predictor (2.3).

**What would falsify a given predictor's claim:** if that field shows no relationship (or the
wrong sign) with next-period deposit growth once size tier is controlled for — in either
transition — that predictor's version of the thesis does not hold up under a real historical
test. Each of the four is falsified independently; one failing does not falsify the others.

## 2.2 Exact data

- **Source table:** `analytics.bank_financial_snapshot` — **not** `bank_financial_snapshot_latest`,
  which is a filtered view holding only the single most recent row per institution and was
  initially (incorrectly) checked first; it looked sparse only because it discards history that
  exists in the base table.
- **Snapshot as of 2026-09-13, banks only:**
  | Period | Institutions |
  |---|---|
  | 2024-12-31 | 4,543 |
  | 2025-03-31 | 4,519 |
  | 2025-09-30 | 4,435 |
  | 2025-12-31 | 4,392 |
  | 2026-03-31 | 4,335 |

  Confirmed this is a real recurring panel, not independent point-in-time samples: of
  institutions present at 2024-12-31, the large majority recur in later quarters (checked
  directly — 22,204 total row-matches against the 2024-12-31 institution list across all
  periods, consistent with most institutions appearing in most of the 5 quarters).

- **Predictor field population, checked directly before writing this bar** (banks, the 4 quarters
  used below): `roa` and `nim` are 100% populated in every quarter (4,543/4,543 etc.);
  `loans_to_deposits_pct` and `brokered_deposits_pct` are ~98.7% populated (e.g. 4,485/4,543 at
  2024-12-31) — no material gaps for any of the four.

- **Two transitions, tested independently and identically, for all four predictors.** Chosen from
  the three viable consecutive-quarter pairs (institution overlap checked for all three: 4,517 /
  4,391 / 4,330) by picking the two that share no quarter, keeping them maximally independent:

  | | Transition C1 | Transition C2 |
  |---|---|---|
  | Prediction period | 2024-12-31 | 2025-09-30 |
  | Outcome period | 2025-03-31 | 2025-12-31 |

  (2025-12-31 → 2026-03-31 was the third viable pair; not used here because it shares its
  predictor quarter with C2's outcome quarter, which would make the two transitions less
  independent of each other. Available as a third confirmatory transition later if wanted, not
  required for this document's bar.)

- **Join key:** `inst_key`, present in both quarters of a given transition for the same
  institution.
- **Predictors, each computed entirely from the prediction-period quarter (no look-ahead),
  identically for both transitions — one of:**
  - `roa`
  - `loans_to_deposits_pct`
  - `brokered_deposits_pct`
  - `nim`

  plus, in every case: `total_assets` from the same quarter, used only for the size-tercile
  control, not as a claim in itself.
- **Outcome, computed from the outcome-period quarter, identical for all four predictors:**
  - `dep_growth` = (`total_deposits`<sub>outcome</sub> − `total_deposits`<sub>prediction</sub>) /
    `total_deposits`<sub>prediction</sub>. Computed directly from raw deposit levels rather than
    using the table's own pre-computed `dep_yoy_pct` field, because that field's reference period
    doesn't necessarily match the specific quarter pair chosen here — this keeps the outcome
    definition transparent and auditable the same way as Test 1's, rather than borrowing a
    pre-aggregated number with a different implicit window.
- **Exclusions, decided here and applied identically to both transitions and all four predictors,
  regardless of result:**
  - `total_deposits`<sub>prediction</sub> `< $10,000,000` excluded — trivial/data-quality-outlier
    institutions; the floor is higher than Test 1's because this is institution-level, not
    branch-level (confirmed real institution deposit levels run from tens of millions upward;
    Trustmark alone was ~$15.2B, spot-checked directly).
  - `total_deposits`<sub>outcome</sub> `< 5% of total_deposits`<sub>prediction</sub> excluded —
    the institution-level equivalent of Test 1's closure exclusion: an acquisition or merger
    causing the institution to stop reporting separately, not organic deposit flight.
  - Institutions with a null value for the predictor field being tested in that variant are
    excluded from that variant only (does not affect the other three predictors' populations).
  - `dep_growth` winsorized to **[−20%, +20%]** before any averaging — a much tighter band than
    Test 1's [−90%, +200%], because this window is ~3 months, not 12; a quarter-over-quarter
    swing beyond ±20% for a going-concern institution of this size is almost certainly a
    data/reporting artifact, not organic movement (medians reported alongside means regardless).
- **Stratification, identical for both transitions and all four predictor variants:** institutions
  split into terciles by `total_assets` and, within each size tercile, terciles by the predictor
  field for that variant — 9 cells per transition per predictor, 18 cells per predictor, 72 cells
  total across all four. Same reasoning as Test 1's stratification: institution size plausibly
  correlates with baseline deposit-growth rate for reasons unrelated to fundamentals (e.g., larger
  institutions growing through branch expansion or M&A rather than organic retention), so a
  population-wide correlation without this control risks the same masking effect Test 1 found on
  9/11.

## 2.3 Success bar — decided now, before Phase 2 runs

**Both transitions must independently satisfy all three criteria below, for a given predictor**,
same discipline as Test 1 (1.3) and for the same reason: a pass on only one transition does not
rule out a quarter-specific artifact, which is the entire reason a second transition exists. No
combined or averaged threshold across the two transitions, and no combined threshold across the
four predictors either — each of the 4 predictors × 2 transitions = 8 individual evaluations is
graded on its own.

For a given predictor and a given transition, the claim is considered **supported** on that
predictor/transition only if *all* of the following hold (`predictor` below stands for whichever
of the four fields is being tested; the monotonicity direction follows 2.1's shared hypothesis —
low ROA/NIM = weak, high LDR/brokered% = weak):

1. **Directional monotonicity:** within at least 2 of the 3 size terciles, median `dep_growth`
   moves in the hypothesized direction as the predictor tercile moves from weak → strong (e.g.
   for ROA: low → mid → high ROA sees increasing median growth; for LDR: low → mid → high LDR
   sees *decreasing* median growth, since high LDR is the "weak" end for that field). A single
   non-monotonic cell in a tercile with fewer than 100 institutions does not break this (smaller
   cell sizes than Test 1's 500-branch threshold, reflecting this test's smaller overall
   population — ~4,300 institutions vs. Test 1's ~76,000 branches).
2. **Practical magnitude:** the spread between the weak-tercile and strong-tercile median growth,
   within at least one size tercile, is **≥ 1 percentage point**. Set lower than Test 1's 3-point
   bar because this is a single quarter's growth, not a full year's — a 1-point quarterly gap is
   roughly comparable in annualized terms and is the smallest spread that would still be a
   practically meaningful signal at this cadence, not a mechanical quartering of Test 1's number.
   Applied identically (same 1-point bar) to all four predictors, since it's a statement about the
   outcome variable's scale, not the predictor's.
3. **Not an artifact of M&A:** the result must hold (same direction, magnitude within 0.5 points)
   after excluding institutions where outcome-period deposits fell below 5% of prediction-period
   deposits (2.2's exclusion) — i.e., the effect is not being driven by acquisitions/mergers
   miscounted as organic flight.

**Multiple-comparisons note, decided now:** testing four predictors instead of one increases the
chance that at least one passes by chance alone, even though each individually still has to clear
both independent transitions (which is itself a real guard against pure noise for any single
predictor). This is handled by reporting, not by picking a stricter per-predictor bar after the
fact:
- Each of the 4 predictors' outcomes (Supported / Not supported / Mixed, per 1.3's definitions,
  applied per predictor across its own 2 transitions) is reported individually — never
  cherry-picked to whichever one happens to pass.
- **All four supported** is treated as strong evidence for the general "institution financial
  health predicts deposit growth" thesis.
- **Exactly one of four supported** is treated as weak/suspect evidence for that one specific
  metric, not evidence for the general thesis — flagged explicitly as such rather than presented
  as "Test 2 passed."
- **Two or three of four supported** falls between these and should be reported with which
  specific predictors passed and which didn't, not summarized as a single pass/fail.

---

# Test 3 — Parent-institution financial health → that specific branch's deposit growth

Test 2 tested the easiest version of the financial-health claim: an institution's own health
predicting its own aggregate growth. That can mask real branch-level heterogeneity (Sprint 6
already found an 11-point score spread across just 4 branches at one institution). Test 3 tests
the actual mechanism `vuln_score` uses in production: a competitor's institution-level financials
predicting one specific branch's deposit growth.

## 3.1 Claim

Same 4 predictors and shared hypothesis direction as Test 2 (2.1), but the unit of analysis for
the *outcome* is the branch, not the institution: a branch's parent-institution ROA/LDR/brokered
%/NIM at time T predicts that specific branch's deposit growth over the following period, after
controlling for the branch's own market share (Test 1's control, since branch-level growth still
has the urbanicity confound Test 1 found — institution size alone doesn't cover it at this grain).

## 3.2 Data

- **Predictor:** `analytics.bank_financial_snapshot`, `institution_type='bank'`, joined to
  `raw.raw_sod` via `RSSDID`.
- **Outcome:** branch-level `yoy_growth`, same definition as Test 1 (1.2).
- **Transition:** predictor quarter **2024-12-31** (earliest available in
  `bank_financial_snapshot`) → outcome from Test 1's existing 2024→2025 branch panel.
  **Known limitation, stated plainly:** 2024-12-31 falls ~6 months into that outcome window
  (which starts at the June 2024 SOD snapshot), not fully before it — there is no Call Report
  quarter earlier than Dec 2024 available yet. This is a partial look-ahead, not a clean
  pre-registration in the strict sense Test 1/2 achieved. Only one transition is possible right
  now for this reason — no second independent quarter exists far enough back. **Single-transition
  results here are directional only, not eligible for a "Supported" verdict** until a second,
  cleaner transition becomes available (either an earlier Call Report quarter gets loaded, or a
  2026 SOD snapshot lands, giving a real post-2024-12-31 outcome window).
- **Exclusions:** Test 1's branch-level floor/closure exclusions (1.2), plus Test 2's null-predictor
  exclusion, applied together.
- **Stratification:** terciles by branch `own_share` (Test 1's control) × terciles by the parent
  institution's predictor value — 9 cells per predictor.

## 3.3 Bar

Same monotonicity + magnitude criteria as Test 1 (1.3: ≥2 of 3 share terciles, ≥3pp spread —
branch-level, so Test 1's bar applies, not Test 2's quarterly-scaled one). **No pass/fail verdict
is issued from this single transition** — report the pattern (direction, magnitude) honestly, and
treat it as informative-only pending a second transition, per 3.2.

---

# Test 4 — Local market tailwind → branch deposit growth

Third mechanism, distinct from competition (Test 1) and financial health (Test 2/3): does the
branch's local market growing (income, population, home values) predict that branch's deposit
growth. Checked data before drafting, same discipline as before:

- `raw.raw_income` — ZCTA-level median household income, YEAR 2021-2024 (real annual panel).
- `raw.raw_population` — ZCTA-level population (`total`), YEAR 2021-2024.
- `raw.raw_zhvi` — ZIP-level home value index, monthly, 2023-01 to 2026-07.

## 4.1 Claim

A branch's local-market growth (income, population, or home-value growth, computed before the
prediction period) predicts that branch's deposit growth over the following period, controlling
for the branch's own market share (same confound as Test 1 — kept for consistency, same
9-cell design).

## 4.2 Data

- **Geography join:** `raw_income`/`raw_population`'s `"Geographic Area Name"` is `"ZCTA5 <zip>"`
  — extract the 5-digit ZIP and join to `raw.raw_sod`'s `ZIPBR`. `raw_zhvi.zip` joins directly.
- **Income and population — two transitions, mirroring Test 1's:**

  | | Transition A | Transition B |
  |---|---|---|
  | Predictor (income/pop growth) | 2021→2022 | 2022→2023 |
  | Branch outcome (SOD) | 2023→2024 | 2024→2025 |

  Predictor year precedes the branch outcome window by design (no look-ahead) — a 1-year gap in
  both cases.
- **ZHVI — single transition only, directional (same limitation class as Test 3):** ZHVI only
  starts Jan 2023, so no ZHVI predictor exists before the 2023→2024 branch transition. Only
  ZHVI growth Jan 2023→Jan 2024 → branch 2024→2025 outcome is possible, and even that predictor
  window ends just months before the branch outcome period starts (same "not fully clean"
  caveat as Test 3.2). **Directional only, no verdict**, same rule as Test 3 (3.3).
- **Predictor/outcome/exclusions/stratification:** identical to Test 1 (1.2) — same floor,
  closure exclusion, winsorization, and `own_share` × predictor-growth tercile grid.

## 4.3 Bar

Income and population: Test 1's bar exactly (1.3) — both transitions must independently pass.
ZHVI: directional only, no verdict, per 4.2.

---

## Phase 2 scope (not started)

Run all four tests as specified above:

- **Test 1:** the query in 1.2 against `raw.raw_sod`, once for each of Transitions A and B,
  producing two 9-cell tercile tables (mean + median `yoy_growth`, n per cell), evaluated
  independently against 1.3.
- **Test 2:** the query in 2.2 against `analytics.bank_financial_snapshot`, once for each of the 4
  predictors (ROA, LDR, brokered deposits %, NIM) × 2 transitions (C1, C2) = 8 runs, each
  producing a 9-cell tercile table (mean + median `dep_growth`, n per cell), evaluated
  independently against 2.3, with the multiple-comparisons reporting discipline from 2.3 applied
  when summarizing across the four predictors.
- **Test 3:** the query in 3.2, once per predictor (4 runs, single transition), reported as
  directional-only per 3.3 — no Supported/Not Supported verdict from this round.
- **Test 4:** the query in 4.2 — income and population, once each × 2 transitions = 4 runs,
  evaluated against 4.3 (Test 1's bar); ZHVI, single transition, directional-only per 4.2.

Commit the queries and all result tables (2 for Test 1, 8 for Test 2, 4 for Test 3, 5 for Test 4 —
19 total) for Phase 3 review. Report each test's outcome separately, and within Tests 2, 3 and 4
report each predictor's outcome separately — none of this is combined into a single overall a1
verdict at this stage.
