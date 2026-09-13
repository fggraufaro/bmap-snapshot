# a1 — Test 1 Results (Phase 2)

Run against `docs/a1_deposit_flight_backtest_methodology.md` section 1.2/1.3, exactly as
pre-registered. Both transitions run fresh on 2026-09-13 — the 2024→2025 numbers are not reused
from the informal 9/11 run (they're close, since 2024/2025 SOD data hasn't materially changed,
but were re-derived independently rather than copied).

Queries run via Supabase MCP (`execute_sql`, project `tuiiywphoynbmkxpoyps`) against
`raw.raw_sod`. Exact query text for each transition (with and without the closure exclusion) is
in `docs/a1_test1_queries.sql` in this same commit.

## Transition A — 2023 → 2024

**With closure exclusion (as specified):**

| share↓ comp→ | 1 (low) | 2 (mid) | 3 (high) |
|---|---|---|---|
| 1 (low share) | n=485, median +0.80% | n=4,795, median +3.55% | n=19,060, median +0.18% |
| 2 (mid share) | n=4,318, median +2.20% | n=15,077, median +0.63% | n=4,944, median **−2.46%** |
| 3 (high share) | n=19,537, median +1.00% | n=4,467, median −0.73% | n=335, median −2.61% |

**Criterion 1 (monotonicity, ≥2 of 3 share terciles):** share-tercile 2 (+2.20 → +0.63 → −2.46)
and share-tercile 3 (+1.00 → −0.73 → −2.61) both monotonically decrease. share-tercile 1 does
not (+0.80 → +3.55 → +0.18). **2 of 3 terciles monotonic → PASS** (met without needing the
<500-branch small-cell exception — noting for the record that share1/comp1's n=485 would have
qualified for that exception anyway, but it isn't needed since share2 and share3 alone clear the
"at least 2 of 3" bar).

**Criterion 2 (≥3pp spread, at least 1 tercile):** share2: 2.20 − (−2.46) = **4.66pp**.
share3: 1.00 − (−2.61) = **3.61pp**. Both exceed the 3pp bar. **PASS.**

**Criterion 3 (robust to closure exclusion, within 1pt):** without the exclusion, share2 spread
= 2.12 − (−2.49) = 4.61pp (Δ0.05pt from 4.66); share3 spread = 1.00 − (−2.61) = 3.61pp (Δ0.00).
Cell-level differences were 0.00–0.08 percentage points throughout, and n changed by only 2 per
cell (almost no branches were actually excluded in this transition). **PASS, clearly.**

**Transition A verdict: all three criteria pass → SUPPORTED.**

## Transition B — 2024 → 2025

**With closure exclusion (as specified):**

| share↓ comp→ | 1 (low) | 2 (mid) | 3 (high) |
|---|---|---|---|
| 1 (low share) | n=531, median +6.25% | n=4,709, median +4.73% | n=18,913, median +1.16% |
| 2 (mid share) | n=4,241, median +4.21% | n=15,014, median +1.36% | n=4,898, median **−0.59%** |
| 3 (high share) | n=19,381, median +2.45% | n=4,430, median +0.95% | n=342, median +2.85% |

**Criterion 1 (monotonicity, ≥2 of 3 share terciles):** share-tercile 1 (+6.25 → +4.73 → +1.16)
and share-tercile 2 (+4.21 → +1.36 → −0.59) both monotonically decrease. share-tercile 3 does
not (+2.45 → +0.95 → +2.85), though its comp3 cell (n=342) is small. **2 of 3 terciles
monotonic → PASS** (again met without needing the small-cell exception).

**Criterion 2 (≥3pp spread, at least 1 tercile):** share1: 6.25 − 1.16 = **5.09pp**.
share2: 4.21 − (−0.59) = **4.80pp**. Both exceed the 3pp bar comfortably. **PASS.**

**Criterion 3 (robust to closure exclusion, within 1pt):** without the exclusion, share1 spread
= 6.07 − 1.15 = 4.92pp (Δ0.17pt); share2 spread = 4.17 − (−0.59) = 4.76pp (Δ0.04pt). Cell-level
differences were 0.00–0.18 percentage points throughout. **PASS, clearly.**

**Transition B verdict: all three criteria pass → SUPPORTED.**

## Overall Test 1 outcome

**Both transitions independently pass all three pre-registered criteria. Test 1 = SUPPORTED.**

Per section 1.3, this is not a coincidence-prone single-period result: the same competitive-density
effect, controlling for market share, replicates across two independent year-pairs (2023→2024
and 2024→2025), with directionally consistent and practically meaningful magnitude (3.6–5.1
percentage points of median deposit-growth spread) in both. The effect also survives removing the
closure exclusion essentially unchanged (all deltas ≤0.18 percentage points), so it is not an
artifact of branch retirements/mergers being miscounted as organic flight.

One honest caveat carried over from the methodology doc, not new here: the population-wide
(unstratified) correlation between competitor density and growth is near zero — this effect only
shows up once market-share tier is controlled for. That is expected and documented in 1.2/1.3,
not a weakness discovered after the fact, but it does mean the sellable claim has to be stated
correctly ("competitive density predicts deposit growth, controlling for market share/size tier")
rather than as a raw, unqualified correlation.

Ready for Phase 3 (Session 4 independent verification).
