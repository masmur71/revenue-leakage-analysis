# Native Excel acceptance — historical checklist

Final acceptance and precise evidence attribution are in [phase4-handoff.md](phase4-handoff.md). Native captures and final operator reports satisfy the Phase 4 checks; final-workbook-integrity.json confirms the macro-free file, embedded model and valid ZIP CRC. The unchecked original execution checklist below is retained as historical procedure, not a current pending-task list. Automated missing-folder failure remained INCONCLUSIVE; final manual recovery and formatting/reopen acceptance are operator-reported, not independently re-observed.

## Capability evidence

Excel Windows is installed. Book1's worksheet/ribbon, ChatGPT button and signed-in pane with normal composer were observed. A connected read returned no queries. The tool contract exposes ordinary range-based PivotTables and Office.js, without native Power Pivot authoring or Save As commands. `typeof` probes returned true for several proxy properties; these do **not** prove supported native operations.

Microsoft documents [Power Query collection read operations](https://learn.microsoft.com/en-us/javascript/api/excel/excel.querycollection?view=excel-js-preview) and states [Office.js does not support OLAP/Power Pivot PivotTables](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-pivottables). The live-control skill limits UI automation to setup/focus and prohibits a COM bypass. After initial manual installation, the user approved a user-run VBA installer for remaining measures only. RevenueLeakage.xlsx exists locally; stages 1–9 are user-reported complete, not independently accepted. The agent has not executed VBA or changed the native model.

## Action checklist

Record Excel version/bitness, workbook path, manifest hash, timestamp, operator and snapshot. Capture screenshots with visible filters and readable values. Do not tick items using CSV-only evidence.

- [ ] New workbook saved as excel/RevenueLeakage.xlsx; rediscover its session instead of reusing Book1's ID.
- [ ] All 14 model tables loaded (6 dimensions, 7 facts, 1 context); SQLReference worksheet-only; no large fact on a sheet.
- [ ] All relationships and row counts match the contract/CSV evidence; no blank orphan members or ambiguous paths.
- [ ] All 46 KPI measures and helpers/diagnostics compile in Power Pivot.
- [ ] Compile/import MeasureInstaller.bas in the separate runner; run on the saved target, review the log/backup, save target, rerun and observe added=0, skipped=66. Confirm new measures on ReportContext and preserved existing helpers; typo extras do not count toward 66.
- [ ] Refresh All completes; inspect query refresh state and refreshed counts, not just a timestamp cell.
- [ ] Copy CSV folder to another local directory, change pFolder, refresh and reconcile; restore original path and refresh again.
- [ ] Missing-folder refresh fails visibly; restore parameter, refresh successfully and leave final workbook healthy.
- [ ] Change year 2025 → 2024 and branch → Medan/B003; observe pivot values and charts change. Restore 2025/all branches and confirm values return.
- [ ] Test all SQLFilterControls.csv rows using their inclusive dates and nonblank dimension IDs. For partial dates use a DimDate[date] Between filter on pControls and clear it afterward. Compare every supplied metric.
- [ ] Test all SQLSupportControls.csv rows at the default period for AR/DSO/visits/RFM. credit_billings_90d is an audit component, not a native-results capture metric.
- [ ] Reconcile all 46 default-period metrics to SQLReference.csv, including true zeros/BLANK. AR stays fixed when only operating dates change.
- [ ] Partial/noncontiguous dates and SKU/outlet/salesman/channel filters yield BLANK target/achievement; full-month branch-only selections retain target.
- [ ] Product selection yields BLANK AR/visits/RFM/cohort; salesman selection yields BLANK RFM/cohort. Clearing filters restores previous values.
- [ ] Check SQLCohort observed cells and RFM segment counts; no future cells as zero and no summed multi-month cohort denominator.
- [ ] Price/net/aging residuals zero for supported contexts; undefined ratios remain BLANK; negative results and >100% return ratios retained.
- [ ] Every sheet visually checked: no ####, formula errors, clipped labels, overlaps or blank charts; primary pivots use ThisWorkbookDataModel.
- [ ] Save, close, reopen .xlsx; Refresh All and repeat Medan/2025 and Jakarta/Q1. Model/measures/slicers persist and still reconcile.

## Numeric capture

CSV verification generates native-results.template.csv. Copy to native-results.csv and fill actual_value with **actual native PivotTable results**, never expected SQL values. Preserve blank for genuine BLANK. Use invariant numeric text without grouping/symbols; 6.98% is a fraction such as 0.0698, not 6.98. Reference/support cases use ReportContext's operating dates; other cases use SQLFilterControls dates. The verifier rejects missing/duplicate/unexpected capture rows.

Run `excel/verify_exports.mjs` using the bundled Node command in docs/excel_model.md. Money/AOV compare after 2-decimal rounding; ratios after 2 percentage decimals (fraction × 10000 rounded); days/productivity 2 decimals; counts exactly. Round half away from zero, matching PostgreSQL/Excel ROUND. NULL must remain NULL. Captured numeric matches alone do not prove refresh/filter actions or provenance.

Write actual actions/results to native-execution.md: workbook/version, each refresh, filter actions and before/after values, save/reopen results, screenshots and failures. Native measures and initial slicer/chart checks have succeeded; full acceptance remains pending. The approved staged runner is documented in ../panduan_vba_dashboard.md and still requires native compile/execution. Phase 4 stays open; no completion commit or phase archive yet.
