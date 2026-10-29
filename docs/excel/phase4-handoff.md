# Phase 4 acceptance and handoff — 2026-10-30

Status: Phase 4 accepted through executed native evidence plus the final operator reports below. Final workbook integrity passed on 2026-10-30: embedded Data Model present, no VBA project, ZIP CRC valid. SHA-256: 72dce81f82b2d1c0aefa77201652ceaf81b33e9d8521552ce45a2bd22a35e047. See final-workbook-integrity.json. Deliverables are ready for the Phase 4 commit and local status handoff.

## Deliverable

`excel/RevenueLeakage.xlsx` contains folder-parameterized Power Query, the embedded Data Model, explicit measures, native PivotTables/charts/slicers and five dashboard/documentation sheets. Large facts are model-only. Separate user-run VBA runners are local helpers, not part of the macro-free deliverable.

CSV grain/types/keys and 28 relationships are documented in excel_model.md and the model contract. Revenue leakage indicators remain separate; no sum of overlapping discounts, price shortfall and returns is represented as recoverable loss. Source bintel remained read-only throughout Phase 4.

## Evidence and attribution

| Check | Evidence / result |
|---|---|
| Measure installation | Native log 20261029-190718: all 66 names/formulas read back, added=0/skipped=66 on rerun |
| Model schema and relationships | User-run Audit passed after correcting the reversed/inactive DimOutlet–FactRFM relationship; 14 model-table counts and 28 directed active relationships checked |
| Default KPI and filter reconciliation | Interrupted run 20261029-225202-verification retains all 282 required main rows: 46 default KPIs, 12 commercial samples and 4 support samples; all reconcile |
| Additional diagnostics | All 131 retained extended values reconcile: guards, residuals, fixed snapshot behavior, five RFM segments and 25 observed cohort cells |
| Remaining cohort samples | Run 20261030-003236-cohort-samples: 10/10 values reconcile, including two 2025 cells, pre-onboarding BLANK and undefined cross-month totals |
| Native refresh | Run 20261030-003422-refresh: 46/46 values reconcile for each of ORIGINAL, ALTERNATE and RESTORED; model counts match and original pFolder restored |
| Query recovery display | Native screenshots show loaded query counts without visible errors after source restoration |
| Slicer synchronization | Native screenshots show Medan monthly chart/table and controls at net sales 862259449.78; national restoration also evidenced |
| Formatting repairs | Earlier agent screenshots found ####, clipped labels and overlapping chart ticks. User subsequently reported repairs complete and declined another computer-use check; final visual acceptance is operator-reported, not independently re-screenshotted |
| Save/reopen and repeated filters | User reported the instructed final sequence safe, including Medan → national with matching figures. This is operator attestation; no new agent-controlled reopen occurred |
| Missing-folder error and recovery | Automated test was INCONCLUSIVE because no synchronous VBA error surfaced. After receiving explicit manual test instructions, user reported all safe. Final manual failure/recovery acceptance is operator-reported; no error screenshot was supplied |

Actual capture files remain labelled `.partial.csv` where the exhaustive macro was interrupted. They were checked against the complete required main key set and completed-case checkpoints; hashes and exact coverage are recorded in partial-capture-review.json. The 276 omitted exhaustive cohort cases are **not** claimed tested. The brief requires reconciled totals and sampled combinations; all required main combinations and representative cohort samples are covered.

Financial displays reconcile after two decimals, ratios after two percentage decimals, counts exactly, and true BLANK separately from zero. Prior numerical evidence was not discarded because the exhaustive run was interrupted.

## Reproduction and operational notes

- Export through `sql/05_excel/02_export.ps1`; generated source CSVs remain local, with hashes in data/excel/manifest.json.
- Follow the Indonesian installation guide for model setup and panduan_penyelesaian_phase4.md for the short remaining-check workflow. Do not rerun the exhaustive Verify suite. Public Phase4_Verify now only warns.
- Preserve final manual formatting. Build/Finish reapply layout and should not be rerun on the accepted workbook merely for validation.
- Validate stored evidence with review_partial_capture.mjs, check_dashboard_capture.mjs --cohort-samples and --refresh; no native macro execution is needed to recompare those saved values.
- Set pFolder to the local CSV directory before refresh on another machine. If a refresh test is interrupted, restore the exact original parameter from its run folder before saving.

## Next phase

Phase 5: write a 1–2 page evidence-based recommendation document covering finding, magnitude, action and explicitly assumed impact, using reproducible SQL references and avoiding overlapping impact estimates. Retain the synthetic-data limitation. Begin in a new session after the Phase 4 commit and local status archive are complete.
