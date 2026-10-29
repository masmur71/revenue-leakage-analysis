# Phase 0 Data Audit

Verified on 2026-10-29 against PostgreSQL 16.15, database `bintel`, schema `analytics`. Source tables were accessed in a read-only transaction. The audit creates a session-local temporary table before that transaction; no source objects are changed.

## Table grain and size

| Table | Grain / primary key | Rows |
|---|---|---:|
| dim_branch | Branch / branch_id | 15 |
| dim_date | Day / date_sk | 731 |
| dim_outlet | Outlet / outlet_id | 3,000 |
| dim_outlet_credit_term | Outlet term version / credit_term_sk | 6,000 |
| dim_product | SKU / sku_id | 300 |
| dim_salesman | Salesman / salesman_id | 50 |
| dim_salesman_assignment | Outlet assignment version / assignment_sk | 3,000 |
| fact_invoices | Invoice / invoice_id; sales_order_id also unique | 129,653 |
| fact_sales | Sales line / sales_line_id | 329,263 |
| fact_returns | Return event / return_id | 6,725 |
| fact_payments | Payment / payment_id | 121,924 |
| fact_visits | Visit / visit_id | 105,918 |
| fact_targets | Target / target_id; branch-month checked unique | 360 |

Total: 706,939 rows. The column/type/nullability inventory is included in both audit outputs; exact DDL is supplied separately. There are 45 declared constraints and 32 existing indexes. Index definitions and validated constraint flags are recorded in `schema-inventory.txt`.

## Relationships and historical joins

Sales lines reference invoices, dates, outlets, branches, products, recorded/assigned salesmen and assignment versions. Returns reference sales lines; payments reference invoices. Outlets reference branches; historical terms and assignments reference outlets. Aggregate child events before joining at invoice or sales-line grain to avoid fan-out.

Credit terms have 3,000 versions for 2024 and 3,000 for 2025. Stored intervals are inclusive: 2024-01-01 through 2024-12-31, then 2025-01-01 through 2025-12-31. Even current rows have finite end dates. Join using the event date within both bounds, not `is_current`. Each invoice matches exactly one historical credit term; each sales line matches exactly one assignment interval. No overlapping intervals were found. Assignment history contains one version per outlet, limiting rotation analysis.

## Verified controls

All 40 audit checks passed: 31 declared foreign-key orphan checks plus nine financial, temporal, uniqueness and visit checks. Primary-key uniqueness is enforced by PostgreSQL; all declared constraints are validated. This is a new project audit, not a reproduction of the 203 checks claimed in the source summary.

| Control total | IDR |
|---|---:|
| Gross sales | 28,960,053,341.78 |
| Discounts | 58,408,378.04 |
| Returns | 363,967,685.02 |
| Net sales | 28,537,677,278.72 |
| Gross profit | 6,493,439,139.72 |
| Payments | 25,504,883,726.22 |

Invoice amounts match summed sales before returns. Return amounts, quantities and returned COGS match sales-line aggregates. Sales arithmetic identities hold exactly. Effective calls have a same-day outlet sale; this does not establish same-salesman attribution.

Sales, payments and visits span 2024-01-01 through 2025-12-31; returns begin 2024-01-02. Targets are text values ranging from `2024-01` through `2025-12`. Every payment has allocation status `Allocated`; alternate allocation behavior cannot be inferred from this dataset.

## Restore verification and remaining decisions

The owner/privilege-free custom dump is 19,642,051 bytes. It restored successfully into new database `leakage_lab`. All 13 row counts, all 40 controls, totals and column metadata in the audit outputs match exactly. This verifies the selected controls, not byte-level equality of every restored row.

The [KPI contract](kpi_dictionary.md) now defines DSO denominator, overdue boundary, treatment of returns in receivables, snapshot timing, effective-call attribution and leakage overlap. Finite SCD intervals require explicit handling outside the observed period. Synthetic data and limited history restrict business generalization. Do not subtract returns again from the already adjusted `fact_sales.net_sales`.
