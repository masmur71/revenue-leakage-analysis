# Native Excel visual review — 2026-10-30

Status: NOT ACCEPTED — formatting repairs required. Inspection used the user-authorized Windows computer-use route. No macro, refresh, filter selection or cell-value/formula edit was performed. Only window size, active sheet/selection, viewport and Executive Summary zoom changed. Excel was left on Methodology.

Final operator update: user subsequently reported formatting repairs completed and declined another computer-use check. Treat the findings below as historical. Final visual acceptance is operator-reported; these screenshots document the earlier defects, not the repaired final workbook. See phase4-handoff.md for attribution.

## Observations

| Sheet | Actual observation | Required correction |
|---|---|---|
| Executive Summary | KPI cards and monthly actual/target chart populated. At 100% the right-side cards/slicers are outside the viewport; at 90% some large cards display ####; 50% reveals the broad layout but is too small for normal reading. Original branch slicer remains apart from the main slicer stack. | Widen or redesign card cells, make a practical readable layout, and provide a chart title. Do not accept reduced zoom alone as a formatting fix. |
| Leakage Diagnostics | Branch reference_sales/price_shortfall display ####; branch and SKU names and some headers are clipped. Price shortfall chart has the expected Medan bar, but full IDR tick labels overlap. Return-reason table is populated and legible. | Choose shared column widths that fit every pivot on the sheet; abbreviate axis units explicitly (for example IDR millions) without changing values; retain distinct leakage indicators. |
| Operations | overdue_ar and RFM monetary display ####. Branch/salesman/segment names and measure headers are clipped. Cohort matrix is populated, with pre-onboarding cells blank; it is wider than the viewport. | Widen shared label/money columns and wrap headers; use readable calendar labels and intentional horizontal navigation for cohort. Check grand-total behavior separately. |
| Data & Calculations | SQLReference table populated; General format displays rounded large numbers and scientific notation for a numerator. Earlier visible pcontrol values match the recorded default totals. | Apply explicit display formats and adequate widths to technical reference columns, preserving stored numeric precision and ratio units. |
| Methodology | All 18 methodology lines visible and legible, including snapshot/period, BLANK, rounding and overlap limitations. | No blocking text-layout issue observed in this view. |

Repeated AutoFit per PivotTable is a likely implementation contributor: later tables share worksheet columns with earlier tables and can narrow columns needed by other pivots. This is an implementation inference, not an additional numerical defect. The prior numeric reconciliations remain valid; visual acceptance remains open.

## Screenshots

All paths are relative to this document's directory:

- `screenshots/native-executive-top.png`
- `screenshots/native-executive-overview.png`
- `screenshots/native-leakage-top.png`
- `screenshots/native-leakage-chart.png`
- `screenshots/native-leakage-sku.png`
- `screenshots/native-leakage-outlet.png`
- `screenshots/native-leakage-returns.png`
- `screenshots/native-operations-ar.png`
- `screenshots/native-operations-visits.png`
- `screenshots/native-operations-rfm.png`
- `screenshots/native-operations-cohort.png`
- `screenshots/native-calculations-top.png`
- `screenshots/native-methodology.png`

These are unaltered native window captures; none is a synthesized dashboard preview. Current inspection does not establish save/reopen or missing-folder failure behavior. Do not close Phase 4 until repairs and remaining native checks pass.
