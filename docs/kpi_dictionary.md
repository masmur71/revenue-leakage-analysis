# KPI Contract — Phase 1

## Reporting conventions

Default operating period: **2025-01-01 through 2025-12-31**, inclusive. Default snapshot: **2025-12-31**, end of day. Parameters are `period_start`, `period_end`, and `as_of_date`; require `2024-01-01 <= period_start <= period_end <= as_of_date <= 2025-12-31`. Dates have no intra-day timestamps. Do not use the machine's current date for historical reporting.

All sources belong to `analytics`. The executable contract is `sql/02_kpi/00_metric_context.sql`; `01_reference_metrics.sql` produces values and `02_validate_metrics.sql` checks data and edge cases. These scripts create only session-local objects. Each `kpi_*` name below denotes a temporary view, not a new public database object.

The metric interface is `(metric_id, numerator, denominator, value, unit)` with `value = numerator / NULLIF(denominator, 0)`. Amount/count metrics use denominator 1. Ratios remain fractions, not percentages; undefined ratios are NULL and display as N/A, not zero. Preserve PostgreSQL numeric precision until display: IDR two decimals, ratios two percentage decimals, days/AOV/productivity two decimals, counts whole numbers. Compute totals from summed components; never average branch percentages. Counts are distinct only when explicitly stated.

SQL reference values are enterprise totals. Future branch/region/SKU/outlet filtering must filter numerator and denominator consistently at the declared grain. AR and RFM are snapshots and must not be summed across snapshots. All monetary amounts are nominal IDR as stored; no unsupported tax, inflation or currency adjustment is added.

## Sales, returns and target achievement

Base sales grain: one `fact_sales.sales_line_id`, filtered by `sales_date` in the operating period. Return grain: one `fact_returns.return_id`, filtered by `return_date` in that period. `kpi_sales`, `kpi_returns`, and `kpi_commercial` implement the definitions below.

| Metric ID | Formula and source | Interpretation |
|---|---|---|
| gross_sales | SUM(fact_sales.gross_sales) | Recorded quantity × pre-discount unit price; not reference-price revenue. |
| discount_amount | SUM(discount_amount) | All recorded sales-line discounts. |
| billed_sales | SUM(net_sales_before_returns) | Gross less discounts, before returns. |
| return_amount | SUM(fact_returns.return_amount) | Merchandise returns on return event dates. |
| net_sales | billed_sales − return_amount | Primary event-period revenue KPI. |
| net_cogs | SUM(fact_sales.cogs_sold) − SUM(fact_returns.returned_cogs) | Event-period cost after returned inventory reversal. |
| gross_profit | net_sales − net_cogs | Gross profit, not operating profit. |
| gross_margin_rate | gross_profit / net_sales | Undefined for zero net sales; negative denominators/results are retained. |
| discount_rate | discount_amount / gross_sales | Weighted monetary rate. |
| return_rate | return_amount / billed_sales | Event-period value ratio; not a probability or sold-cohort return rate; may exceed 100%. |
| sales_units | SUM(fact_sales.quantity) | Gross sold units, not net of returns. |
| returned_units | SUM(fact_returns.return_quantity) | Returned units by event date. |
| return_margin_reversal | return_amount − returned_cogs | Signed gross-profit reversal; do not confuse return revenue with lost margin. |
| active_outlets | COUNT(DISTINCT fact_sales.outlet_id) | At least one recorded sale in the period; full later returns do not erase activity. |
| orders | COUNT(DISTINCT sales_order_id) | Recorded sales orders; multiple lines count once. |
| average_order_value | billed_sales / orders | Sales-date order value before returns, avoiding returns on prior-period orders. |
| target_net_sales | SUM(fact_targets.target_net_sales) | Branch-month targets; parse target_month as YYYY-MM-01 and validate against dim_date. |
| achievement_rate | net_sales / target_net_sales | Period actual event revenue versus matching branch-month target. |

Targets are comparable only for complete calendar months: target and achievement return NULL for partial-month boundaries. There is no proration and no SKU/salesman target allocation. Data checks require one target per branch-month. Empty sums of transactions are zero; a missing target remains missing and fails coverage validation.

`fact_sales.return_amount`, `returned_cogs`, `net_sales`, `net_cogs`, and `gross_profit` include linked returns regardless of the report cutoff. They can support a separately labeled **sale-cohort restatement**, but must not be used as event-period/as-of amounts. Never subtract returns twice. Whole-dataset event and restated totals agree, while months and years can differ. A later return is recognized in its return month even if this makes that month's net sales negative.

## Leakage indicators and non-overlapping bridge

All price/discount metrics below use sales-line grain and `sales_date`, implemented in `kpi_sales_totals`. The reference price is the supplied baseline, not a demonstrated collectible selling price.

| Metric ID | Formula | Interpretation |
|---|---|---|
| reference_sales | SUM(quantity × reference_price) | Reference-price gross value. |
| price_shortfall | SUM(GREATEST(quantity × reference_price − gross_sales, 0)) | Below-reference price exposure before explicit discounts. |
| price_upside | SUM(GREATEST(gross_sales − quantity × reference_price, 0)) | Above-reference price value. |
| price_deviation_rate | (gross_sales − reference_sales) / reference_sales | Signed, quantity/reference-weighted price deviation. |
| invalid_promo_discount | SUM(discount_amount) WHERE promotion_id IS NOT NULL AND is_promo_valid = false | Discount exposure on explicitly invalid attached promotions. |
| invalid_promo_line_rate | Invalid attached-promo lines / all attached-promo lines | Non-promo lines excluded from both components. |
| salesman_mismatch_sales | SUM(net_sales_before_returns) WHERE is_salesman_match = false | Sales recorded by a different salesman from the assigned one; not a financial loss. |
| salesman_mismatch_rate | Mismatched lines / all sales lines | Line-weighted incidence, not order/outlet incidence. |

Bridge: **reference_sales + price_upside − price_shortfall − discount_amount = billed_sales**. For event-period net revenue subtract event-period return_amount afterward. Invalid-promo discounts are already inside discount_amount and must never be added again. Return margin reversal already affects gross profit. Do not sum these indicators into a single recoverable-loss number.

Observed in the full dataset: 329,211 non-promo rows have `is_promo_valid=false`, while 52 attached-promo rows are valid. Thus the invalid-promo KPI is zero, not a reason to treat every non-promo discount as improper. NULL flags are unknown; this dataset declares the flag NOT NULL. `promo_discount_pct` is not substituted for actual recorded discounts.

Recovery estimates belong to Phase 5 and require explicit scenarios (eligible exposure × assumed recovery fraction), not a claim that all discounts, returns or reference-price differences are preventable.

## Receivables, collections and DSO

Base grain: one `fact_invoices.invoice_id` issued on/before the snapshot. `kpi_invoice_snapshot` aggregates payments first by invoice, including only `allocation_status='Allocated'` and `payment_date <= as_of_date`. Payments after cutoff and unallocated payments do not settle AR. Any previously unseen allocation state is flagged for review by validation rather than silently accepted as settled.

Historical terms use `dim_outlet_credit_term` with `invoice_date >= valid_from AND (valid_to IS NULL OR invoice_date <= valid_to)`, exactly one match. Credit invoices use that historical term. Cash invoices use 0 days even if the outlet has a credit agreement. Contractual due date is invoice_date + applicable term, reconciled to stored due_date. Missing/overlapping history fails validation; never fall back silently to current terms. Due on the snapshot is **not overdue**.

Invoice balance = invoice_amount − allocated_paid. Positive balance is outstanding; a negative balance is separately reported as overpayment, never used to offset another invoice's debt. There are no credit-note/refund records proving that returns settle invoices. Consequently no merchandise-return deduction is applied to AR; this is billed receivables, with return-related settlement uncertainty disclosed. It is not a legal debt determination.

| Metric ID | Formula in kpi_ar / kpi_metric_values | Grain / timing |
|---|---|---|
| outstanding_ar | SUM(GREATEST(invoice balance, 0)) | All eligible invoices, including cash residuals, at snapshot. |
| overdue_ar | Outstanding where as_of_date − due_date > 0 | Strictly after contractual due date. |
| overdue_rate | overdue_ar / outstanding_ar | Value-weighted, all billed AR. |
| ar_not_due | Outstanding with days past due <= 0 | Includes due today and future due dates. |
| ar_1_30 | Outstanding with 1–30 days past due | Inclusive bounds. |
| ar_31_60 | Outstanding with 31–60 days past due | Inclusive bounds. |
| ar_61_90 | Outstanding with 61–90 days past due | Inclusive bounds. |
| ar_over_90 | Outstanding with >90 days past due | Starts at day 91. |
| credit_ar | Outstanding on payment_method='Credit' | Credit-only DSO numerator. |
| cash_ar | Outstanding on payment_method='Cash' | Separate exception monitor. |
| overpayment | SUM(GREATEST(allocated_paid − invoice_amount, 0)) | Not netted against outstanding. |
| allocated_receipts | SUM(fact_payments.payment_amount) with Allocated status in operating period | Cash receipts by payment date, not snapshot/cumulative collections. |
| dso_90d | credit_ar × 90 / credit invoice_amount in [as_of_date−89, as_of_date] | Trailing 90 calendar days; gross billed credit sales, before returns; not lifetime sales or monthly averages. |

DSO is NULL if 90 days of source coverage are unavailable or the denominator is zero. It can exceed 90 days. It estimates collection velocity; do not interpret it as average days past due. Do not call payment receipts divided by period sales a cohort collection rate. Invoice, payment and return timestamps differ.

## Salesman productivity and order conversion

Source `fact_visits`, filtered by visit_date. Raw visit grain is visit_id; conversion grain is **(salesman_id, visit_date, outlet_id)** to prevent multiple lines/orders or duplicate visits from multiplying successes. Attribution requires the same recorded salesman, date and outlet in fact_sales; assigned_salesman_id is not substituted. Sales need not be linked to a visit because no visit/order foreign key exists; same-day conversion is association, not causal proof.

| Metric ID | Formula in kpi_visit_totals | Interpretation |
|---|---|---|
| scheduled_visits | COUNT(*) | All logged visit rows, including cancellations; not unrecorded plans. |
| completed_visits | COUNT(*) WHERE is_completed | Completed visit events. |
| effective_calls | COUNT(*) WHERE is_completed AND is_effective_call | Recorded effective visit events; validator requires an attributed same-day sale. |
| visit_completion_rate | completed_visits / scheduled_visits | Event-grain completion. |
| effective_call_rate | effective_calls / completed_visits | Uses completed visits, not scheduled visits. |
| order_conversion_rate | Distinct completed calls with attributed orders / distinct completed calls | Deduplicated call-grain conversion. |
| calls_per_salesman_day | completed_visits / COUNT(DISTINCT (salesman_id, visit_date)) among completed visits | Productivity per observed active field day, not headcount, hire date or working-calendar capacity. |

Do not average individual salesman rates for enterprise totals. Flag mismatch uses sales_line_id grain and the recorded-versus-assigned IDs; visit conversion uses the recorded salesman. No statement about coverage of all planned working days is supported.

## RFM and onboarding retention

`kpi_rfm` produces one row per outlet onboarded on/before snapshot. Window: **365 calendar days ending at snapshot**, inclusive `[as_of_date−364, as_of_date]`. For early snapshots, history is truncated at 2024-01-01 and must be labeled partial; do not compare its scores with a complete window. No imputation for pre-2024 transactions.

| Output | Formula and source | Edge behavior |
|---|---|---|
| recency_days | as_of_date − MAX(fact_sales.sales_date in window) | NULL if no order in window. |
| frequency | COUNT(DISTINCT sales_order_id in window), per outlet | Zero for inactive outlets; returned orders still count. |
| monetary | SUM(net_sales_before_returns in window) − SUM(return_amount by return date in window), per outlet | Event revenue; may be negative, including returns of older orders. |
| r_score | Recency <=30: 5; <=90: 4; <=180: 3; <=270: 2; else: 1 | Zero when frequency=0. |
| f_score | CEIL(5 × CUME_DIST() OVER ORDER BY frequency ASC) among active outlets | 1–5; equal frequencies share a score; some bands may be empty. |
| m_score | CEIL(5 × CUME_DIST() OVER ORDER BY monetary ASC) among active outlets | Same tie rule; negative monetary retained; inactive score=0. |

`kpi_rfm_scored` implements all three scores. Higher means more recent/frequent/higher monetary value. Scores are relative to each snapshot's active population for F/M; do not equate a score change with absolute growth. Keep R/F/M separately; business segment labels and actions belong to Phase 2.

`kpi_cohort` outputs cohort_month, activity_month, month_age, cohort_outlets, active_outlets and retention_rate. Cohort is onboarding month from dim_outlet, not first-purchase month. Denominator is all outlets onboarded in that month, including never-purchasers. Numerator is distinct cohort outlets with at least one recorded sale in the activity month. Rate = numerator/denominator; returns do not remove activity. Month zero is same-month activation, not repeat retention.

Only cohorts onboarded from 2024-01-01 onward are compared; earlier cohorts are left-censored. Only complete observed calendar months are emitted. Generate the cohort/month grid so zero-sales cells remain zero, and never emit future/unobserved cells as zero. Month age is calendar-month difference; denominators remain fixed across observed months. Missing months with no observations are not retention failures.

## Validation, reference values and limits

`02_validate_metrics.sql` validates all input domains, historical term coverage, due dates, financial/price bridges, aging partition, attributed calls, target coverage, RFM reconciliation/scores, cohort bounds, and zero denominators. Explicit fixtures exercise late payments, unallocated payments, aging boundaries, SCD year changes, later-period returns and false flags without promotions. A failed check raises an exception and a nonzero psql exit code.

The reference script emits all 46 scalar metric IDs plus RFM control totals, score distribution and cohort cells. Per-outlet RFM values are available through `kpi_rfm_scored` in the same session. Phase 2 will add drilldowns and business interpretations without silently changing this contract.

Source data is synthetic, observation ends 2025-12-31, payment status only contains Allocated, assignments have no rotation, and promotions are sparse. The model does not establish credit-note balances, causal effects, legal collectibility, reference-price realizability or unobserved planned visits. Volume/price/discount variance against target requires additional baseline components: only branch-month monetary targets exist. Do not fabricate target quantities or prices.
