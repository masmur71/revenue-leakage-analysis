# Excel model and dashboard — Phase 4

For step-by-step installation in Indonesian, see [Panduan pemasangan Excel](panduan_pemasangan_excel.md).

**IN PROGRESS. No native workbook has been built or accepted yet.** CSV/M/DAX files are implementation inputs, not proof of working native Excel features. The user selected a new workbook, all KPI areas with leakage priority, and manual M/DAX installation when native authoring is unavailable.

## Reproduce inputs

Run from the repository root in PowerShell with Docker running:

```powershell
./sql/05_excel/02_export.ps1
$node = 'C:\Users\masmu\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe'
& $node excel/build_contract.mjs
& $node excel/verify_exports.mjs
```

`bintel` stays read-only. The export creates session-local views before a REPEATABLE READ READ ONLY transaction; no source data, permanent schema or indexes change. Original Phase 1/2 definitions remain authoritative. Phase 3's distinct-call/aggregate-first approaches inform the export without changing metrics.

`data/excel/manifest.json` records CSV hashes and historical parameters. Full transaction history through `as_of_date` is exported so period slicers work. `period_start`/`period_end` define default acceptance, not transaction truncation. Defaults: 2025-01-01–2025-12-31, snapshot 2025-12-31. Snapshot changes require re-export/refresh; operating-period slicers never change AR/RFM snapshots. Never use TODAY(). Generated M resets the folder parameter to this checkout; adjust it for another machine.

Do not refresh Excel while regenerating CSVs. A failed export is not accepted: resolve the log failure, rerun successfully and validate before refresh. Large CSVs are local reproducible outputs excluded from Git; their manifest and verification evidence remain reviewable.

## Dataset contract

CSV: UTF-8, comma delimiter, quoted fields, header, empty field = NULL, PostgreSQL booleans `t`/`f`. Dates: ISO `YYYY-MM-DD`; month keys use day 01. Numeric money has decimal `.`, no symbols/grouping, and exact source cents. IDs stay text. `excel/model_contract.json` specifies every column and primary key: money → `Currency.Type`, integers → `Int64.Type`, dates → Date, ratios/control values → decimal number. `fnLoadCsv` rejects schema/type/key errors instead of replacing them with zero.

| Dataset | Grain / unique key | Purpose |
|---|---|---|
| DimBranch | branch / branch_id | Region attributes included |
| DimOutlet | outlet / outlet_id | No branch relationship; prevents ambiguous paths |
| DimProduct | SKU / sku_id | SKU and category |
| DimSalesman | recorded salesman / salesman_id | Home branch descriptive only |
| DimChannel | channel / channel | Shared channel filter |
| DimDate | calendar day / date | Full observed calendar, including zero-activity dates |
| FactCommercial | event / event_id (`S:` sales_line_id or `R:` return_id) | Independent sale/return dates; no lifetime restatement columns |
| FactVisits | visit / visit_id | Also salesman/date/outlet call_key and salesman/date salesman_day_key |
| FactInvoiceSnapshot | invoice at one snapshot / invoice_id | Payments aggregated before join; invoice-date SCD terms |
| FactTargets | branch-month / target_id | Additional unique branch_id + month_start |
| FactPayments | payment / payment_id | Receipt-period amounts, separate from invoice settlement |
| FactRFM | outlet at one snapshot / outlet_id | Fixed national SQL scores, no rescoring after slicers |
| FactCohort | outlet × activity month / cohort_cell_id | Fixed onboarding denominators; never-purchasers retained |
| ReportContext | exactly one row | Parameters/completeness |
| SQLReference | metric / metric_id | 46 independent SQL controls; worksheet only |

`SQLBranchMonth`, `SQLYoY`, `SQLCohort`, `SQLFilterControls`, `SQLSupportControls` are reference/acceptance exports, not additive facts. Never union diagnostic totals across branch/SKU/outlet cuts. SQLYoY describes observed change, not target causality. Large facts must load into Data Model only.

## Manual native installation

For the staged user-run dashboard automation, see [the Indonesian dashboard runner guide](panduan_vba_dashboard.md). `excel/DashboardRunner.bas` preserves unambiguous existing objects, builds missing native model pivots/charts/slicers, captures actual values on an isolated probe, and tests refresh with parameter recovery. Native execution and reconciliation evidence is summarized in [the acceptance record](excel/phase4-handoff.md). Generate reviewed input specs with `node excel/build_dashboard_specs.mjs`; do not regenerate against changed exports without rerunning source validation. Use the short completion guide rather than restarting exhaustive Verify, and preserve final manual formatting.

1. Create a **new** Excel Windows workbook; save as `excel/RevenueLeakage.xlsx`. Keep Book1 separate. Open ChatGPT in the new workbook for later connected-session verification. Do not change macro/security settings.
2. Data → Get Data → From Other Sources → Blank Query → Advanced Editor. Paste `excel/power_query/pFolder.m`, name it `pFolder`; point it to this checkout's absolute `data/excel` folder. Create `fnLoadCsv` from its `.m` file. Both are connection-only.
3. Create one blank query for each remaining `.m` file using the exact filename stem as query name. Dim*, Fact*, ReportContext: Close & Load To → **Only Create Connection** + **Add this data to the Data Model**. SQLReference: worksheet table at `Data & Calculations!A5`, not the model. Do not auto-combine different schemas or apply automatic type detection.
4. Create all relationships in `excel/relationships.csv`: dimension key = one side, fact key = many side. No dimension-to-dimension/fact-to-fact links. Do not relate DimDate to InvoiceSnapshot, RFM or Cohort; their date semantics differ.
5. Power Pivot → Manage: mark DimDate as date table using `date`, confirm types, then install the 66 definitions in `excel/measures.dax` using the user-run `excel/MeasureInstaller.bas` and [Indonesian instructions](panduan_vba_measures.md). New measures use ReportContext as home to avoid FactCommercial column/name collisions; existing matching helpers are preserved on their current home table. Manual alternative: `name := expression` maps to Measure Name = name and Formula = `=expression`, in dependency order. Hide `_` helpers and raw money columns from client tools after verification.
6. IDR/AOV format `#,##0.00`; rates `0.00%`; days/productivity `0.00`; counts `#,##0`; dates `yyyy-mm-dd`. Preserve BLANK, with visible legend **Blank = N/A / unsupported filter**. Ratios remain fractions.
7. Create sheets/pivots below via Insert → PivotTable → **This Workbook's Data Model**. Cards/charts read native pivots via GETPIVOTDATA or PivotCharts, never copied SQL values. Save only after successful refresh.

### Dashboard specification

Tab order: Executive Summary, Leakage Diagnostics, Operations, Data & Calculations, Methodology. English labels; white background, dark navy headers, teal revenue, restrained orange/red for shortfalls. Keep all monetary displays in explicit IDR with 2 decimals. Readable widths, no clipped labels or merged data ranges. Reserve A:L for content, N:U for slicers. Row 1 holds title/period/snapshot; Executive Summary A2:H5 holds cards. Leave sufficient space for pivot refresh expansion.

| Location | Pivot | Fields / measures |
|---|---|---|
| Data & Calculations!J5 | pExecutive | Values net_sales, gross_profit, gross_margin_rate, target_net_sales, achievement_rate, price_shortfall, discount_amount, return_amount; executive cards link here |
| Executive Summary!A10 | pTrend | Rows DimDate[year_month]; values net_sales, target_net_sales; actual vs target chart at A40 |
| Executive Summary!F10 | pBranch | Rows DimBranch[region_name], branch_name; values net_sales, target_gap, achievement_rate |
| Leakage Diagnostics!A10 | pLeakBranch | Rows DimBranch[branch_name]; values reference_sales, price_upside, price_shortfall, discount_amount, billed_sales, return_amount, net_sales, gross_profit |
| Leakage Diagnostics!A35 | pLeakSKU | Rows DimProduct[sku_name]; values price_shortfall, return_amount, return_rate, billed_sales; top 10 by return_amount descending |
| Leakage Diagnostics!G35 | pLeakOutlet | Rows DimOutlet[outlet_name]; same values/sort; top 10 by return_amount |
| Leakage Diagnostics!A60 | pReasons | Rows FactCommercial[return_reason], exclude blank; values return_amount, return_margin_reversal |
| Data & Calculations!J30 | pControls | All 46 KPI values, price_bridge_residual, net_bridge_residual, aging_residual; compare with SQLReference |
| Operations!A10 | pAR | Rows DimBranch[branch_name]; values outstanding_ar, overdue_ar, overdue_rate, dso_90d |
| Operations!G10 | pVisits | Rows DimSalesman[salesman_name]; completed_visits, visit_completion_rate, effective_call_rate, order_conversion_rate, calls_per_salesman_day |
| Operations!A80 | pRFM | Rows FactRFM[segment]; values rfm_outlets, rfm_monetary |
| Operations!A110 | pCohort | Rows FactCohort[cohort_month], columns activity_month; values retention_rate; disable grand totals across activity months |

Keep unfiltered detail available when inspecting ties or small denominators. Move the Values field to Rows in pExecutive and pControls for narrow two-column audit blocks. If another pivot grows beyond its reserved block, move the adjacent pivot before adding charts; do not overlap objects or shrink fonts. The cohort heatmap occupies A110:AA160; the area above remains A:L content plus slicers.

Create dimension slicers: DimDate[year], DimDate[year_month], DimBranch[region_name], DimBranch[branch_name], DimProduct[category], DimProduct[sku_name], DimOutlet[outlet_id], DimSalesman[salesman_name], DimChannel[channel]. Place in N2:U32 on the two main sheets. Report Connections must link them to all relevant model pivots. Copy connected slicers to the other sheet and verify synchronized state. Default saved view: clear all filters, then year 2025.

Date slicers affect operating facts only. AR/RFM labels explicitly say **Snapshot: 2025-12-31, independent of operating period**. Cohort uses its own activity-month columns. Product filters make AR/visits blank; product/salesman filters make RFM/cohort blank. Targets/achievement are blank for SKU/outlet/salesman/channel filters and partial/noncontiguous periods. Branch/region full-month targets remain valid. Do not create a shared return-reason slicer. Cohort totals across several activity months intentionally return BLANK; month zero is activation, not repeat retention.

### Methodology content

Reference sales + price upside − price shortfall − all discounts = billed sales; subtract return-event revenue once for net sales. Invalid attached-promo discounts are a subset of discounts. Salesman mismatch is an exception exposure, not proven loss. Return revenue and COGS reversal use return dates; lifetime restatement columns are excluded. Do not report one summed recoverable leakage number; recovery scenarios belong to Phase 5.

AR is billed receivables: merchandise returns are not proven credit notes. Only allocated payments known at snapshot settle invoices. Positive balances and overpayments remain separate; due today is not overdue. Historical invoice-date SCD terms apply. Ratios use summed components. Conversion uses unique salesman/date/outlet calls; productivity uses distinct observed salesman-days. RFM is a trailing 365-day event window. Cohorts include never-purchasers and only complete observed months. Synthetic data does not establish causality, reference-price collectibility or unlogged working capacity.

Place source links to docs/kpi_dictionary.md, docs/core_analysis.md, docs/optimization_report.md, sql/05_excel/00_export.sql and manifest hashes/parameters on Methodology.

## Completion gate

Follow `docs/excel/native-verification.md`. CSV verification is independent integer-cent/rational arithmetic, **not** M/DAX execution. A successful save or matching cached numbers does not prove refresh/filter behavior. Only after native checks pass: inspect Git status, stage explicit deliverables excluding local-only paths, commit, archive STATE.md as archieve/STATE-phase-4.md, and create a carry-over STATE.md of at most 10 lines. Until then keep Phase 4 open and do not create a completion commit.
