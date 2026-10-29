# Interrupted native verification review — 2026-10-30

The existing user-run capture was reviewed without opening Excel or executing VBA. Raw `.partial.csv` files were preserved under `excel/dashboard-runs/20261029-225202-verification/`; they were not renamed to imply full-run completion.

## Verified results

| Evidence | Result |
|---|---|
| Main capture | All 282 required rows present, unique, matched to completed CASE_CAPTURED checkpoints and reconciled |
| Default KPIs | All 46 reconcile with SQLReference |
| Commercial samples | All 12 cases × 15 metrics reconcile with SQLFilterControls |
| Supporting samples | All 4 cases × 14 metrics reconcile with SQLSupportControls |
| Saved extended capture | All 131 rows reconcile; no duplicates, unknown keys, mismatches or uncommitted cases |
| Guards | Product/outlet/salesman/channel, partial and noncontiguous dates produce expected BLANK; full-month branch target retained |
| Residuals/snapshots | Saved price/net residual cases and aging/snapshot-year check reconcile |
| RFM | All five saved segment count/monetary checks reconcile with the validated snapshot controls |
| Cohort | 25 observed cells reconcile: January 2024 cohort across 24 activity months, plus February 2024 age zero |

Money/AOV and days/productivity match at two decimals, ratios at two percentage display decimals, counts exactly, and BLANK separately from zero. `verify_exports.mjs` also reran the independent source validation successfully. Its 860 comparisons include source comparisons plus numeric main-capture comparisons; all 282 main rows were checked, including BLANKs.

## Minimum remaining work

Do **not** rerun `Phase4_Verify`. The interrupted exhaustive suite still has 276 cohort-related cases (826 values), but the brief requires reconciled totals and sampled combinations, not exhaustive evaluation of every cohort cell in one long macro. Those omitted cases remain explicitly unverified.

To complete representative cohort coverage, inspect only two additional observed cells (2025-01 cohort at 2025-01, and 2025-12 cohort at 2025-12), one structurally unobserved cell (2025-12 cohort at 2025-01, expected BLANK), absence of future activity columns, and disabled/undefined cross-month totals. The original `cohort_future` case actually uses a pre-onboarding activity month; do not describe it as a post-snapshot future date.

Remaining independent acceptance: original/alternate/missing-folder refresh and recovery with numeric controls; visual QA of all dashboard sheets; slicer/chart synchronization and reset including year changes; save/close/reopen, refresh and repeat Medan/2025 and Jakarta/Q1. Keep previously recorded successful evidence. No new Excel operation was performed during this review.

## Reproduction

From the project root, using the bundled Node executable:

```powershell
& $nodePath excel/verify_exports.mjs --native excel/dashboard-runs/20261029-225202-verification/native-results.partial.csv
& $nodePath excel/review_partial_capture.mjs excel/dashboard-runs/20261029-225202-verification
```

Machine-readable results and SHA-256 hashes are in `partial-capture-review.json`; source/main verification is in `csv-verification.json`. The full macro did not finish. Phase 4 remains open; no completion commit or archive is justified yet.
