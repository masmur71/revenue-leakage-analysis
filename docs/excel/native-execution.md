# Native Excel execution record

Status: Phase 4 acceptance checks satisfied using native captures and final operator reports; final workbook integrity passed. See phase4-handoff.md and final-workbook-integrity.json. Earlier pending/failure notes below are chronological history, superseded where explicitly resolved in the handoff record.

Computer-use visual inspection: selected the unique RevenueLeakage - Excel window, maximized it and navigated Executive Summary to A1. Native screenshot showed populated KPI cards and monthly actual/target chart; right-side cards/slicers extend beyond viewport at 92/100% zoom. Opening Zoom caused the helper to return only a cropped root-window image and disabled accessibility root; the Zoom modal was not separately targetable in list_windows. Rehydration and runtime reset did not recover a usable modal capture, so further input stopped. User must close Zoom before screenshot QA can continue. No data/filter changes, macros or refresh were executed. Technical evidence: screenshots/native-ui-zoom-capture-blocked.png. This is a capture blocker, not visual acceptance.

Restored query-state screenshots: screenshots/stage12-refresh-restored-query-facts.png and stage12-refresh-restored-query-bottom.png show loaded counts without visible errors for the fact tables, ReportContext and SQLReference. Earlier user screenshot shows dimensions loaded and local pFolder text. This supports healthy displayed state after recovery; it does not establish that the missing-folder refresh visibly failed. Extra relationships query shows 29 loaded rows; this is query metadata, not a claim that the model has 29 relationships (model audit passed the 28-relationship contract).

Native results 2026-10-30: cohort run 20261030-003236-cohort-samples reconciles all 10 values, including two 2025 observed cells, pre-onboarding BLANK and multi-month BLANK totals. Refresh run 20261030-003422-refresh reconciles 46 reference values each for ORIGINAL, ALTERNATE and RESTORED (138 values), with expected model counts. Original pFolder was restored by the logged recovery path. Missing-folder refresh returned no synchronous VBA error, so the runner correctly stopped as INCONCLUSIVE after recovery. Visible query-state/failure evidence, visual QA and save/reopen remain pending; no full macro rerun is needed.

Update 2026-10-30: user reports Audit passed. Verification run 20261029-225202-verification was interrupted after prolonged cohort processing. Existing partial files contain all 282 required main values and 131 extended values with completed-case checkpoints; all reconcile. This includes all 46 default KPIs, 12 commercial and 4 support cases, guards/residuals, all five RFM segments and 25 cohort cells. See partial-verification-review.md and partial-capture-review.json for precise coverage, hashes and minimum remaining acceptance. Full Verify must not be restarted merely to repeat these successful captures. No Excel or VBA execution occurred during this review; full suite and phase acceptance remain incomplete.

Staged dashboard automation is implemented in excel/DashboardRunner.bas and documented in docs/panduan_vba_dashboard.md. Native progress and the current failure are recorded below. Its input package defines 13 dashboard/support pivots, 343 verification cases and 1239 actual-value requests. Offline helper/contract tests are preparation evidence only; no runner-generated native numeric capture exists yet.

Update 2026-10-29: user-run native dashboard logs confirm all 13 pivots created/reused, including the empty branch pivot adopted in place. Relationship orientation/inactive RFM outlet issue was corrected by the user and subsequent model preflight passed. Build remains blocked at slicer view creation, error 5. Run 20261029-222659 records valid level [DimDate].[year].[year]; level-only change did not resolve it. Current code connects pivots before view creation and separates minimal creation from layout; Phase4_FinishDashboard resumes this stage. This revision still requires native execution. No completed dashboard or numeric capture is claimed.

Update from native run 20261029-223224-dashboard: slicer creation now succeeds through outlet, with 14 connected pivots logged for outlet. Next failure is layout protection: salesman destination overlaps outlet_id. Fixed-row placement has been replaced with measured slicer bottoms plus a 24-point gap and free-floating placement for the N-column slicer stack. Existing views are preserved; this layout revision requires native retry. Full dashboard completion and numeric verification remain pending.

Workbook: excel/RevenueLeakage.xlsx exists on disk; native model contents not independently verified.
Excel version/bitness: not recorded yet.
Manifest SHA-256: not recorded yet.
Operator and execution time: not recorded yet.

## Completed prerequisites

- Source export completed through read-only PostgreSQL transactions.
- CSV/source reconciliation evidence: csv-verification.json.
- Excel Windows and signed-in ChatGPT pane observed in Book1; this is not the new target workbook.
- User initially selected manual M/DAX installation, then approved a user-run VBA installer for remaining measures only.
- Installer preparation passed 268 pure-core assertions in Windows Script Host with mock collections/schema; those automated tests did not open Excel.
- User-run native log excel/measure-install-logs/20261029-190718.csv reports DONE: expected=66, added=0, skipped=66, total native measures=66. Screenshot agrees. All expected names/formulas were read back; rerun added no duplicates.

## Required actual evidence

Record each action with its actual before/after result, manifest hash and screenshot path. Follow native-verification.md; do not copy source values as native results.

| Action | Actual result | Evidence |
|---|---|---|
| Create/save new target workbook | User-reported complete; file exists | Native save/reopen acceptance pending |
| Install typed queries and model relationships | Stages 1–9 user-reported complete | Native inspection pending |
| Run installer and repeat with zero additions | Passed name/formula readback and idempotent rerun; no separate compile screenshot recorded | excel/measure-install-logs/20261029-190718.csv and user screenshot |
| Initial PivotTable numeric results | Three initial measures match SQL at year=2025; remaining measures and slicers pending | screenshots/stage11-2025-initial-controls.png |
| Refresh original folder | User reports Refresh All completed without error; screenshot after report retains matching 2025 values. Query completion state not independently inspected | screenshots/stage11-2025-refresh-reported.png |
| Refresh alternate folder, then restore | Pending | None |
| Missing-folder failure and recovery | Pending | None |
| Default total and filter captures | Initial 2025 net_sales=15762528589.51, gross_profit=3851071463.51, price_bridge_residual=0.00 match; branch/date filter action tests still pending | screenshots/stage11-2025-initial-controls.png |
| Medan/2025 branch filter | Matches SQLFilterControls medan_2025: net_sales=862259449.78, gross_profit=212059567.78, price_shortfall=3023132.22; price_bridge_residual=0.00. Branch filtered in Rows; restoration to all branches not yet evidenced | screenshots/stage11-2025-medan-controls.png |
| Clear branch filter and restore national totals | Passed from screenshot: all 15 branches visible, year=2025, net_sales=15762528589.51, gross_profit=3851071463.51, price_bridge_residual=0.00, price_shortfall=3023132.22 | screenshots/stage11-2025-branch-filter-restored.png |
| Default supporting KPI values | Screenshot matches rounded achievement=100.13%, gross margin=24.43%, outstanding AR=3396761237.52 and DSO=87.26. Target amount matches but caption Sum of target_net_sales suggests implicit column sum; replace/confirm explicit ReportContext target measure. Refresh completion not reported | screenshots/stage11-2025-support-controls-implicit-target.png |
| Cross-sheet slicer synchronization and chart response | User reports branch selection works after connecting pControls. Screenshot confirms monthly table/chart changed under multi-select: net_sales=9579086378.09 and target=9572786329.07. Full selection and pControls not shown; single-branch reconciliation and restoration pending | screenshots/stage12-branch-slicer-chart-response.png |
| Single Medan slicer, chart and controls | Passed from two screenshots: Medan alone selected; monthly table/chart and pControls both net_sales=862259449.78 (matches SQL medan_2025), target=883076716.87. Controls profit=212059567.78 and shortfall=3023132.22 match SQL. Other displayed values: achievement=97.64%, margin=24.59%, AR=186456535.77, DSO=90.23; not separately reconciled here. Clear-slicer restoration pending | screenshots/stage12-medan-slicer-chart.png; screenshots/stage12-medan-slicer-controls.png |
| Clear SKU filter and restore supporting totals | Passed from screenshot: SKU=All, year=2025; net_sales=15762528589.51, gross_profit=3851071463.51, price_shortfall=3023132.22, achievement=100.13%, margin=24.43%, AR=3396761237.52, DSO=87.26, explicit target=15742179409.15. User reports Ctrl+S; save/reopen still pending | screenshots/stage11-2025-sku-cleared-restored.png |
| Product filter guards, year=2025 and SKU004 | Passed from screenshot: target_net_sales, achievement_rate, outstanding_ar and dso_90d are BLANK. Explicit target behavior confirmed. Visible net_sales=38083464.33, gross_profit=10856792.33, price_shortfall=216971.82, margin=28.51%; SKU004 amounts not independently SQL-reconciled yet. Clear-filter restoration pending | screenshots/stage11-2025-sku004-guards.png |
| Visual inspection of every sheet | Pending | None |
| Save/close/reopen/refresh and repeat filters | Pending | None |

Phase 4 exit criteria are not met. No phase-completion commit or archive has been made.
