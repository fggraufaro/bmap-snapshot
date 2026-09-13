# a1 — Test 3 Results (Phase 2, directional only)

Per `docs/a1_deposit_flight_backtest_methodology.md` section 3.3: single transition only
(parent-institution 2024-12-31 Call Report → branch-level 2024→2025 SOD growth), so **no
Supported/Not Supported verdict is issued** — reporting pattern and magnitude only. Full queries
in `docs/a1_test3_queries.sql`.

| Predictor | Direction (low→high, all 3 share terciles) | Pattern | Spreads |
|---|---|---|---|
| ROA | Wrong sign (growth *decreases* as ROA increases) | Clean, consistent | e.g. share1: 2.27%→1.60%→1.38% |
| LDR | Wrong sign (growth *increases* as LDR increases) | Clean, consistent — matches Test 2's LDR finding | e.g. share1: 0.29%→1.99%→4.52% (+4.23pp) |
| Brokered % | Right sign (growth decreases as brokered% increases) | Not monotonic — bumps up at middle tercile in all 3 | -0.47pp to -1.53pp |
| NIM | Right sign (growth increases as NIM increases) | Not monotonic — U-shaped, dips at middle tercile in all 3 | +1.17pp to +1.62pp |

**Notable:** LDR's wrong-sign result replicates Test 2's institution-level finding exactly, now at
branch level too — this looks like a real, direction-consistent pattern across both units of
analysis, not a fluke of aggregation. ROA is also wrong-signed here, more cleanly than at Test 2's
institution level. Brokered % and NIM point the hypothesized direction but neither is cleanly
monotonic (both show a same-shaped anomaly: the *middle* predictor tercile deviates, not the low
or high ends) — worth a second transition before reading much into that shape.

**Next step:** this stays directional until a second transition is available — either an earlier
Call Report quarter gets loaded, or a 2026 SOD snapshot lands giving a real post-Dec-2024 outcome
window. Not proposing further action now; flagging for whenever more data exists.
