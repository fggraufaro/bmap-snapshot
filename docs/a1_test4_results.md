# a1 — Test 4 Results (Phase 2)

Per `docs/a1_deposit_flight_backtest_methodology.md` sections 4.2/4.3. Queries in
`docs/a1_test4_queries.sql`.

## Income — NOT SUPPORTED

| | Transition A (2021→22 → 2023→24) | Transition B (2022→23 → 2024→25) |
|---|---|---|
| share1 | 1.02→0.90→0.35 (wrong sign) | 1.98→1.78→1.43 (wrong sign) |
| share2 | 0.33→0.37→−0.01 (flat) | 1.35→1.21→1.35 (flat) |
| share3 | 0.83→0.44→0.65 (non-monotonic) | 2.28→1.91→2.23 (non-monotonic) |
| Max \|spread\| | 0.67pp | 0.55pp |

Wrong-signed or flat in every tercile, both transitions, magnitudes far under the 3pp bar.
**Not supported.**

## Population — NOT SUPPORTED (magnitude), but most directionally consistent predictor tested so far

| | Transition A | Transition B |
|---|---|---|
| share1 | 0.23→0.81→1.43 (monotonic, +1.20pp) | 1.28→1.79→2.28 (monotonic, +1.00pp) |
| share2 | −0.30→0.01→0.92 (monotonic, +1.22pp) | 0.97→1.25→1.63 (monotonic, +0.66pp) |
| share3 | 0.51→0.73→0.72 (flat) | 2.32→1.77→2.30 (non-monotonic, ~flat) |

Right direction, cleanly monotonic in 2 of 3 share terciles in **both** transitions — the cleanest
directional replication of any predictor tested across all of a1 (Tests 2/3 included). But the
magnitude never exceeds ~1.2pp in either transition, well short of the 3pp bar. **Not supported
by the pre-registered bar** — same no-partial-credit discipline as everywhere else — but flagged
as the most promising non-Test-1 candidate seen, worth a second look if the bar is ever revisited.

## ZHVI — directional only (single transition, no verdict per 4.2)

| share tercile | low→mid→high ZHVI growth | spread |
|---|---|---|
| 1 | 1.47%→1.76%→1.90% | +0.43pp |
| 2 | 0.80%→1.58%→1.42% (non-monotonic) | +0.62pp |
| 3 | 1.89%→2.28%→2.27% (flat at top) | +0.38pp |

Right direction in all 3 terciles, modest magnitude, not cleanly monotonic. No verdict — single
transition only, per the methodology doc's own rule (same as Test 3).

## Overall Test 4 outcome

**Income:** not supported. **Population:** not supported by magnitude, but the most consistent
directional signal found in this entire exercise outside Test 1. **ZHVI:** directional only, right
sign, no verdict. None of Test 4's predictors clear the bar to enter a2's evidence section as-is.

Ready for Phase 3.
