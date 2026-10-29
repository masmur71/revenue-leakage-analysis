# Sales Performance & Revenue Leakage Analysis

Where should a distributor look first when sales miss expectations: prices, returns, or unpaid invoices?

This project uses PostgreSQL and Excel to work through that question with a synthetic FMCG distribution dataset covering 2024–2025. The workbook brings sales, targets, returns and receivables into one place. SQL supplies the calculations and the checks behind them.

## What the data shows

For 2025, net sales were **IDR 15.76 billion**, with a **24.43% gross margin**.

| Area | Finding | What it suggests |
|---|---|---|
| Pricing | IDR 3.02 million below reference prices, about **0.02% of net sales**, all in Medan | A small exception to review, rather than the main business problem |
| Returns | IDR 210.28 million, or **1.32% of billed sales** | Investigate high-value cases and order errors before proposing broad changes |
| Receivables | IDR 3.24 billion overdue, **95.38% of outstanding billed AR** | Start with invoice reconciliation and collection priorities |
| Branch performance | Bandung reached **96.67% of target**, despite sales growing year on year | National growth can hide local shortfalls |

The pricing result is small. This dataset does **not** support a large revenue-leakage story. The stronger investigation is receivables, although the data has no credit-note records to confirm how returns settle invoices.

The [recommendations](docs/recommendations.md) explore possible actions. For scale, preventing 5% of returns would preserve about IDR 10.51 million in annual revenue—just **0.07% of net sales**. That is a scenario to test, not a measured saving.

## Dashboard preview

[Download the workbook](excel/RevenueLeakage.xlsx) to explore branches, months, products and outlets.

![Monthly sales and targets filtered to Medan](docs/excel/screenshots/sales-dashboard.png)

## Run it locally

The README and screenshots are enough for a first look. To inspect the calculations:

1. Restore the bundled database using [setup](docs/setup.md).
2. Run `./sql/05_excel/02_export.ps1` to create the CSV inputs.
3. Open the workbook, set `pFolder` to the CSV directory, and refresh.

The full workbook needs Excel desktop for Windows with Power Query and Data Model support. Database setup uses Docker and PowerShell. [Refresh instructions](docs/refresh.md) cover the workbook settings.

## Behind the analysis

- [Business analysis](docs/core_analysis.md) and [KPI definitions](docs/kpi_dictionary.md)
- [Data dictionary](docs/data_dictionary.md) and [ERD](docs/erd.md)
- [Query optimization](docs/optimization_report.md), including before/after plans and timings

The SQL includes historical credit-term joins so older invoices use the terms valid when issued, rankings to locate price exceptions, and separate return events to keep monthly revenue consistent. Query rewrites were checked for equal results before comparing performance.

This is a practice dataset, not evidence about a real distributor. Its value is in the calculations, SQL choices and how the results are interpreted. It cannot establish which interventions would work in a real business.
