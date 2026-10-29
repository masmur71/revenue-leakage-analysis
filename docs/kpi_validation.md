# Phase 1 Verification

Verified on 2026-10-29 against source `bintel` and restored copy `leakage_lab`, PostgreSQL 16.15 in `bintel-pg`. Reproduce with `sql/02_kpi/03_verify_phase1.ps1`; definitions are in `kpi_dictionary.md`.

## Executed acceptance checks

| Scenario | Operating period | Snapshot | Result |
|---|---|---|---|
| 2025 default | 2025-01-01 – 2025-12-31 | 2025-12-31 | 38 checks passed |
| 2024 historical | 2024-01-01 – 2024-12-31 | 2024-12-31 | 38 checks passed |
| Mid-2025 | 2025-01-01 – 2025-06-30 | 2025-06-30 | 38 checks passed |
| Partial early period | 2024-01-15 – 2024-01-31 | 2024-01-31 | 38 checks passed |
| Entire observed history | 2024-01-01 – 2025-12-31 | 2025-12-31 | 38 checks passed; Phase 0 financial totals match |

The five runs are saved as `kpi-validation-2025.txt`, `kpi-validation-2024.txt`, `kpi-validation-mid2025.txt`, `kpi-validation-partial2024.txt`, and `kpi-validation-full.txt`. The 38 checks are repeated across scenarios, not 190 distinct rules.

All 46 default scalar metrics, RFM controls/score distribution and 300 observed cohort-month cells match exactly between source and restored reference outputs. Four invalid parameter cases were rejected as expected: future snapshot, period ending after snapshot, reversed period and start before source coverage. PowerShell runner syntax was parsed successfully.

## Default reference values

Operating-period values below cover calendar 2025; balance values are snapshots at 2025-12-31. These are definition/control references, not final business recommendations.

| KPI | Value |
|---|---:|
| Event-period net sales | IDR 15,762,528,589.51 |
| Gross profit | IDR 3,851,071,463.51 |
| Gross margin | 24.43% |
| Target achievement | 100.13% |
| Outstanding billed AR | IDR 3,396,761,237.52 |
| Overdue billed AR | IDR 3,239,849,366.28 |
| Overdue rate | 95.38% |
| Credit-billing DSO, trailing 90 days | 87.26 days |
| Effective call rate, completed-visit denominator | 6.80% |
| Deduplicated completed-call order conversion | 6.98% |
| Active outlets | 2,976 |
| Invalid attached-promo discount | IDR 0.00 |

## Timing reconciliation and interpretation

Stored `fact_sales.net_sales` for sale dates in 2025 totals IDR 15,772,251,357.76. It excludes IDR 9,722,768.25 of returns occurring in 2025 against sales from 2024. The event-period contract recognizes those returns in 2025, producing IDR 15,762,528,589.51. `04_revenue_timing_bridge.sql` makes the reconciliation reproducible. Differences from older summary figures must first be checked for period, grain and timing definitions rather than treated as query errors.

The partial-period test confirms target achievement is undefined without full-month boundaries and DSO is undefined without 90 observed days. RFM history completeness is exposed in reference parameters. No inference is made about unobserved transactions, missing credit notes or unknown allocation statuses.

## Acceptance and next phase

Phase 1 exit criteria are met: all specified KPIs have formulas, source/grain/timing definitions and executable references/validations. No permanent source objects or data were changed. Phase 2 should use this contract for diagnostic drilldowns; changes to definitions require updated documentation and verification evidence.
