# Analytics data dictionary

Source: owner-free [schema DDL](../sql/00_setup/analytics_schema.sql), bundled synthetic snapshot. All 13 tables belong to `analytics`. Exact types/nullability below follow DDL; business formulas remain in the [KPI dictionary](kpi_dictionary.md). Monetary fields are nominal IDR; quantities are recorded units. NOT NULL is a storage constraint, not proof of business validity.

## dim_branch

One branch; geographic and region attributes. Primary key: `branch_id`.

| Column | PostgreSQL type | Nullable | Reference |
|---|---|---|---|
| branch_id | text | No | — |
| branch_name | text | No | — |
| region_id | text | No | — |
| region_name | text | No | — |
| city | text | Yes | — |
| province | text | Yes | — |

## dim_date

One calendar date; date_sk is the date join key. Primary key: `date_sk`.

| Column | PostgreSQL type | Nullable | Reference |
|---|---|---|---|
| date_sk | integer | No | — |
| date | date | No | — |
| year | integer | No | — |
| quarter | integer | No | — |
| month | integer | No | — |
| month_name | text | No | — |
| week_of_year | integer | No | — |
| day_of_month | integer | No | — |
| day_of_week | text | No | — |
| is_weekend | boolean | No | — |

## dim_outlet

One outlet; onboarding date, channel and branch attributes. Primary key: `outlet_id`.

| Column | PostgreSQL type | Nullable | Reference |
|---|---|---|---|
| outlet_id | text | No | — |
| outlet_name | text | No | — |
| branch_id | text | No | `dim_branch.branch_id` |
| channel | text | No | — |
| onboard_date | date | No | — |
| initial_credit_term_days | integer | No | — |

## dim_outlet_credit_term

One dated outlet credit-term version; inclusive valid_from/valid_to. Use invoice date, not current flag. Primary key: `credit_term_sk`.

| Column | PostgreSQL type | Nullable | Reference |
|---|---|---|---|
| credit_term_sk | bigint | No | — |
| outlet_id | text | No | `dim_outlet.outlet_id` |
| valid_from | date | No | — |
| valid_to | date | No | — |
| credit_term_days | integer | No | — |
| version | integer | No | — |
| is_current | boolean | No | — |

## dim_product

One SKU; product/category and supplied reference/cost attributes. Primary key: `sku_id`.

| Column | PostgreSQL type | Nullable | Reference |
|---|---|---|---|
| sku_id | text | No | — |
| sku_name | text | No | — |
| category | text | No | — |
| standard_cost | numeric(18,2) | No | — |

## dim_salesman

One salesman; home branch is an attribute, not transaction attribution. Primary key: `salesman_id`.

| Column | PostgreSQL type | Nullable | Reference |
|---|---|---|---|
| salesman_id | text | No | — |
| salesman_name | text | No | — |
| home_branch_id | text | No | `dim_branch.branch_id` |
| hire_date | date | No | — |
| status | text | No | — |

## dim_salesman_assignment

One dated outlet-to-salesman assignment version; inclusive intervals. Only one version per outlet in this snapshot. Primary key: `assignment_sk`.

| Column | PostgreSQL type | Nullable | Reference |
|---|---|---|---|
| assignment_sk | bigint | No | — |
| outlet_id | text | No | `dim_outlet.outlet_id` |
| salesman_id | text | No | `dim_salesman.salesman_id` |
| branch_id | text | No | `dim_branch.branch_id` |
| valid_from | date | No | — |
| valid_to | date | No | — |
| version | integer | No | — |
| is_current | boolean | No | — |

## fact_invoices

One invoice; sales_order_id is also unique. Billed amount before merchandise returns. Primary key: `invoice_id`.

| Column | PostgreSQL type | Nullable | Reference |
|---|---|---|---|
| invoice_id | text | No | — |
| sales_order_id | text | No | — |
| invoice_date | date | No | — |
| due_date | date | No | — |
| outlet_id | text | No | `dim_outlet.outlet_id` |
| branch_id | text | No | `dim_branch.branch_id` |
| salesman_id | text | No | `dim_salesman.salesman_id` |
| channel | text | No | — |
| invoice_amount | numeric(18,2) | No | — |
| credit_term_days | integer | No | — |
| payment_method | text | No | — |
| date_sk | integer | No | `dim_date.date_sk` |

## fact_payments

One payment allocation event. Only Allocated payments at cutoff settle an invoice. Primary key: `payment_id`.

| Column | PostgreSQL type | Nullable | Reference |
|---|---|---|---|
| payment_id | text | No | — |
| invoice_id | text | No | `fact_invoices.invoice_id` |
| payment_date | date | No | — |
| payment_amount | numeric(18,2) | No | — |
| payment_method | text | No | — |
| allocation_status | text | No | — |
| date_sk | integer | No | `dim_date.date_sk` |

## fact_returns

One return event, linked to a sales line; recognize at return_date. Primary key: `return_id`.

| Column | PostgreSQL type | Nullable | Reference |
|---|---|---|---|
| return_id | text | No | — |
| sales_line_id | text | No | `fact_sales.sales_line_id` |
| sales_order_id | text | No | — |
| return_date | date | No | — |
| outlet_id | text | No | `dim_outlet.outlet_id` |
| branch_id | text | No | `dim_branch.branch_id` |
| sku_id | text | No | `dim_product.sku_id` |
| return_quantity | integer | No | — |
| return_amount | numeric(18,2) | No | — |
| returned_cogs | numeric(18,2) | No | — |
| reason | text | No | — |
| date_sk | integer | No | `dim_date.date_sk` |

## fact_sales

One sales line. Sale-date gross/discount/billed amounts; lifetime linked-return columns are NOT event-period KPIs. Primary key: `sales_line_id`.

| Column | PostgreSQL type | Nullable | Reference |
|---|---|---|---|
| sales_line_id | text | No | — |
| sales_order_id | text | No | — |
| invoice_id | text | No | `fact_invoices.invoice_id` |
| sales_date | date | No | — |
| date_sk | integer | No | `dim_date.date_sk` |
| outlet_id | text | No | `dim_outlet.outlet_id` |
| branch_id | text | No | `dim_branch.branch_id` |
| salesman_id | text | No | `dim_salesman.salesman_id` |
| assigned_salesman_id | text | No | `dim_salesman.salesman_id` |
| assignment_sk | bigint | No | `dim_salesman_assignment.assignment_sk` |
| sku_id | text | No | `dim_product.sku_id` |
| channel | text | No | — |
| quantity | integer | No | — |
| unit_price_before_discount | numeric(18,2) | No | — |
| gross_sales | numeric(18,2) | No | — |
| discount_amount | numeric(18,2) | No | — |
| net_sales_before_returns | numeric(18,2) | No | — |
| return_quantity | integer | No | — |
| return_amount | numeric(18,2) | No | — |
| returned_cogs | numeric(18,2) | No | — |
| net_sales | numeric(18,2) | No | — |
| unit_cost | numeric(18,2) | No | — |
| cogs_sold | numeric(18,2) | No | — |
| net_cogs | numeric(18,2) | No | — |
| gross_profit | numeric(18,2) | No | — |
| gross_margin_pct | numeric(18,6) | Yes | — |
| reference_price | numeric(18,2) | No | — |
| price_deviation_pct | numeric(18,6) | Yes | — |
| promotion_id | text | Yes | — |
| promo_discount_pct | numeric(18,4) | Yes | — |
| is_promo_valid | boolean | No | — |
| is_returned | boolean | No | — |
| is_salesman_match | boolean | No | — |

## fact_targets

One branch/month target (validated uniqueness); target_month is YYYY-MM text. No volume/price/discount target components. Primary key: `target_id`.

| Column | PostgreSQL type | Nullable | Reference |
|---|---|---|---|
| target_id | text | No | — |
| branch_id | text | No | `dim_branch.branch_id` |
| target_month | text | No | — |
| target_net_sales | numeric(18,2) | No | — |
| currency | text | No | — |
| date_sk | integer | No | `dim_date.date_sk` |

## fact_visits

One visit event. Conversion deduplicates salesman/date/outlet; same-day matching is association. Primary key: `visit_id`.

| Column | PostgreSQL type | Nullable | Reference |
|---|---|---|---|
| visit_id | text | No | — |
| visit_date | date | No | — |
| date_sk | integer | No | `dim_date.date_sk` |
| outlet_id | text | No | `dim_outlet.outlet_id` |
| branch_id | text | No | `dim_branch.branch_id` |
| salesman_id | text | No | `dim_salesman.salesman_id` |
| visit_status | text | No | — |
| is_completed | boolean | No | — |
| is_effective_call | boolean | No | — |

## Interpretation guards

- `fact_sales.net_sales`, `return_amount`, `returned_cogs`, `net_cogs`, and `gross_profit` are lifetime sale-cohort restatements. Use the KPI event-date views for period results; never deduct returns twice.
- `reference_price` is a supplied comparison baseline, not proven collectible revenue. `is_promo_valid=false` only indicates invalid attached promotion when `promotion_id` is present.
- `assigned_salesman_id` differs conceptually from recorded `salesman_id`; the mismatch flag is an incidence indicator, not a monetary loss.
- Invoice credit terms require exactly one inclusive historical match; Cash uses zero days. AR includes cash residuals and excludes unallocated/future payments, without inferred return credit notes.
- Synthetic IDs and contact-like fields are fabricated dataset attributes. Date coverage ends on 2025-12-31; see [audit](data_audit.md) for counts, coverage and validation.
- Derived CSV grains, data types, keys and 28 workbook relationships are in the [Excel model guide](excel_model.md) and [model contract](../excel/model_contract.json). They are not additional source tables.
