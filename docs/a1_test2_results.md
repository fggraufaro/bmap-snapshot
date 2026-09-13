# a1 — Test 2 Results (Phase 2)

Run against `docs/a1_deposit_flight_backtest_methodology.md` sections 2.2/2.3, exactly as
pre-registered. All 8 evaluations (4 predictors × 2 transitions) run 2026-09-13. Full query text
in `docs/a1_test2_queries.sql`.

Criterion-3 (M&A-exclusion robustness) checks were skipped for predictor/transition pairs that
already failed criteria 1 or 2 with the exclusion applied — a candidate that fails the main test
doesn't need a robustness check, and this is noted per-predictor below rather than silently
omitted.

## ROA — NOT SUPPORTED

| | C1 (2024-12-31→2025-03-31) | C2 (2025-09-30→2025-12-31) |
|---|---|---|
| Direction | Increasing in 2-3 size terciles (flattens at top tercile in 2 of 3) | Increasing in 2 of 3 (dips in size2) |
| Max spread (any size tercile) | 0.76pp | 0.97pp |
| Clears 1pp bar? | No | No (closest of the two, but still under) |

Both transitions land under the 1pp magnitude bar. Direction is broadly right but the effect is
too small to count as practically meaningful under the bar set in advance. Criterion-3 check
skipped (already fails on magnitude).

## Loans-to-Deposits Ratio — NOT SUPPORTED (wrong sign, replicated)

| | C1 | C2 |
|---|---|---|
| Direction | Increasing (opposite of hypothesized) in all 3 size terciles | Increasing (opposite) in all 3, less cleanly monotonic |
| Spread, low→high LDR | size1 +1.75pp, size2 +0.92pp, size3 +1.28pp | size1 +0.83pp, size2 +0.71pp, size3 +0.15pp |

The hypothesis was: higher LDR = weaker funding position = worse subsequent growth. The data show
the opposite, cleanly and consistently, in both independent transitions: higher-LDR banks grew
deposits *faster*, not slower. This isn't a "no effect" result — it's a real, replicated,
wrong-signed effect. Per section 2.1's falsification criterion, this directional hypothesis is
falsified, not just unsupported. (Plausible untested explanation: banks running higher LDR may be
actively, successfully competing for deposits to fund loan demand rather than being passively
squeezed — but that's a new hypothesis for a future test, not something this document tests.)
Criterion-3 check skipped (already fails on direction).

## Brokered Deposits % — NOT SUPPORTED

| | C1 | C2 |
|---|---|---|
| Spread, low→high brokered% | size1 +0.04pp, size2 −0.07pp, size3 +0.74pp | size1 +0.21pp, size2 +0.32pp, size3 +0.23pp |

No tercile in either transition shows the hypothesized decreasing pattern with meaningful
magnitude — results are near-zero or mildly wrong-signed throughout, not monotonic in either
direction. This was flagged in the methodology doc as the mechanistically sharpest of the four
candidates; empirically here it's the weakest. Criterion-3 check skipped (already fails on
direction/magnitude).

## NIM — MIXED → NOT SUPPORTED (closest of the four)

| | C1 | C2 |
|---|---|---|
| Direction | Monotonically increasing, all 3 size terciles | Monotonically increasing, all 3 size terciles |
| Max spread | **1.37pp** (size1) | 0.84pp (size1) |
| Clears 1pp bar? | **Yes** | No |
| Robust to M&A exclusion? | Yes — identical with/without (0 institutions excluded this transition) | Not checked (already fails magnitude) |

NIM is the only predictor that's directionally clean in *both* transitions (monotonic increase,
right sign, every size tercile, both periods) — but it only clears the 1pp magnitude bar in one
of the two (C1: 1.37pp; C2: 0.84pp, close but short). Per section 2.3's explicit rule, a
one-transition pass is graded as **Mixed → Not Supported**, the same discipline applied to Test 1
throughout — no exception made here despite how close C2 came.

## Overall Test 2 outcome

**0 of 4 predictors supported.** Per section 2.3's multiple-comparisons reporting discipline:

- **ROA, LDR, Brokered %:** cleanly not supported — wrong sign (LDR), no pattern (Brokered %), or
  right direction but insufficient magnitude in both transitions (ROA).
- **NIM:** the closest result — right direction in both transitions, passes in one, narrowly
  misses in the other. Worth flagging as the strongest candidate for a future look (e.g., against
  the third available quarter pair, 2025-12-31→2026-03-31, not used here since it shares a quarter
  with C2 — see methodology doc section 2.2), but not something to claim as validated now.

**This means the "institution financial health predicts deposit growth" thesis, in the specific
form pre-registered here, is not supported by this data.** That's a real, informative result, not
a failure of the exercise — it means Test 1's branch-density signal should not be assumed to
generalize to institution-level financial fundamentals without further, separately-scoped work.
Nothing here should be used in the BlastPoint pitch or a2's evidence section as institution-level
predictive evidence; Test 1's branch-density result stands on its own and is unaffected by this.

Ready for Phase 3 (Session 4 independent verification).
