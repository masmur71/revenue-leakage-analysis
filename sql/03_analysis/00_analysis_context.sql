-- Phase 2 shared diagnostics. Sources: analytics via the Phase 1 KPI contract.
-- Session-local views only. Create temporary metadata before the read-only report transaction (PostgreSQL restriction).
\ir ../02_kpi/00_metric_context.sql

-- Business intent: locate price, discount and return exposure without overlap.
-- Grain: one event row, then one (dimension, member). Returns use return dates.
-- No sales/returns fact-to-fact join; UNION ALL preserves independent event grain.
CREATE TEMP VIEW analysis_events AS
SELECT sales_date event_date,branch_id,outlet_id,sku_id,quantity units,
 quantity*reference_price reference_sales,gross_sales,
 greatest(quantity*reference_price-gross_sales,0) price_shortfall,
 greatest(gross_sales-quantity*reference_price,0) price_upside,
 discount_amount discounts,net_sales_before_returns billed_sales,cogs_sold,
 CASE WHEN promotion_id IS NOT NULL AND NOT is_promo_valid THEN discount_amount ELSE 0 END invalid_promo_discount,
 0::numeric returns,0::numeric returned_cogs,0::bigint returned_units
FROM kpi_sales
UNION ALL
SELECT return_date,branch_id,outlet_id,sku_id,0,0,0,0,0,0,0,0,0,
 return_amount,returned_cogs,return_quantity FROM kpi_returns;
CREATE TEMP VIEW analysis_leakage AS
WITH cuts AS (
 SELECT 'branch' dimension,branch_id member_id,e.* FROM analysis_events e
 UNION ALL SELECT 'sku',sku_id,e.* FROM analysis_events e
 UNION ALL SELECT 'outlet',outlet_id,e.* FROM analysis_events e
), totals AS (
 SELECT dimension,member_id,sum(units) units,sum(reference_sales) reference_sales,
 sum(gross_sales) gross_sales,sum(price_shortfall) price_shortfall,sum(price_upside) price_upside,
 sum(discounts) discounts,sum(billed_sales) billed_sales,sum(cogs_sold) cogs_sold,
 sum(invalid_promo_discount) invalid_promo_discount,sum(returns) returns,
 sum(returned_cogs) returned_cogs,sum(returned_units) returned_units
 FROM cuts GROUP BY 1,2
)
SELECT *,billed_sales-returns net_sales,billed_sales-returns-cogs_sold+returned_cogs gross_profit,
 (billed_sales-returns-cogs_sold+returned_cogs)/nullif(billed_sales-returns,0) margin_rate,
 discounts/nullif(gross_sales,0) discount_rate,returns/nullif(billed_sales,0) return_rate,
 returns-returned_cogs return_margin_reversal,
 dense_rank() OVER(PARTITION BY dimension ORDER BY price_shortfall DESC) price_exposure_rank,
 sum(price_shortfall) OVER(PARTITION BY dimension ORDER BY price_shortfall DESC,member_id ROWS UNBOUNDED PRECEDING)
 /nullif(sum(price_shortfall) OVER(PARTITION BY dimension),0) cumulative_price_share
FROM totals;

-- Business intent: branch/region monthly achievement and explanatory components.
-- Grain: all branch-months intersecting period, including zero-transaction months.
-- Targets only apply when the entire calendar month is inside the report period.
-- Validation: additive actuals reconcile to KPI; target is joined after aggregation.
CREATE TEMP VIEW analysis_branch_month AS
WITH actual AS (
 SELECT branch_id,date_trunc('month',event_date)::date AS month,
 sum(units) units,sum(gross_sales) gross_sales,sum(discounts) discounts,
 sum(billed_sales) billed_sales,sum(returns) returns,
 sum(billed_sales-returns-cogs_sold+returned_cogs) gross_profit
 FROM analysis_events GROUP BY 1,2
), grid AS (
 SELECT b.*,m::date AS month,p.period_start,p.period_end FROM analytics.dim_branch b
 CROSS JOIN kpi_params p CROSS JOIN LATERAL generate_series(date_trunc('month',p.period_start),date_trunc('month',p.period_end),interval '1 month') m
), joined AS (
 SELECT g.branch_id,g.branch_name,g.region_id,g.region_name,g.month,
 coalesce(a.units,0) units,coalesce(a.gross_sales,0) gross_sales,
 coalesce(a.discounts,0) discounts,coalesce(a.billed_sales,0) billed_sales,
 coalesce(a.returns,0) returns,coalesce(a.billed_sales,0)-coalesce(a.returns,0) net_sales,
 coalesce(a.gross_profit,0) gross_profit,
 g.month>=g.period_start AND (g.month+interval '1 month' - interval '1 day')::date<=g.period_end complete_month,
 CASE WHEN g.month>=g.period_start AND (g.month+interval '1 month' - interval '1 day')::date<=g.period_end THEN t.target_net_sales END target
 FROM grid g LEFT JOIN actual a ON a.branch_id=g.branch_id AND a.month=g.month
 LEFT JOIN analytics.fact_targets t ON t.branch_id=g.branch_id AND t.target_month=to_char(g.month,'YYYY-MM')
)
SELECT *,net_sales-target target_gap,net_sales/nullif(target,0) achievement_rate,
 lag(net_sales) OVER(PARTITION BY branch_id ORDER BY month) previous_observed_month_sales
FROM joined;

-- Descriptive prior-year bridge; NOT a target variance or causal attribution.
-- Quantity and blended price include SKU mix; there are no SKU quantity targets.
-- Exact sequential bridge: (Q1-Q0)*P0 + Q1*(P1-P0) - change(discounts) - change(returns).
-- Grain: branch/month in report; prior-year sales use matching clipped date bounds.
CREATE TEMP VIEW analysis_yoy AS
WITH historical_events AS (
 SELECT sales_date d,branch_id,quantity q,gross_sales g,discount_amount discount,0::numeric r FROM analytics.fact_sales
 UNION ALL SELECT return_date,branch_id,0,0,0,return_amount FROM analytics.fact_returns
), prior AS (
 SELECT branch_id,date_trunc('month',d+interval '1 year')::date AS month,sum(q) units,
 sum(g) gross_sales,sum(discount) discounts,sum(r) returns
 FROM historical_events CROSS JOIN kpi_params p
 WHERE d BETWEEN (p.period_start-interval '1 year')::date AND (p.period_end-interval '1 year')::date GROUP BY 1,2
)
SELECT a.branch_id,a.month,a.net_sales,p.gross_sales-p.discounts-p.returns prior_net_sales,
 a.net_sales-(p.gross_sales-p.discounts-p.returns) net_change,
 (a.units-p.units)*p.gross_sales/nullif(p.units,0) volume_at_prior_blended_price,
 a.units*(a.gross_sales/nullif(a.units,0)-p.gross_sales/nullif(p.units,0)) blended_price_mix_change,
 -(a.discounts-p.discounts) discount_change,-(a.returns-p.returns) return_change
FROM analysis_branch_month a LEFT JOIN prior p USING(branch_id,month);

-- Return reason grain: branch/outlet/SKU/reason; amount and margin are distinct.
CREATE TEMP VIEW analysis_return_reasons AS
SELECT branch_id,outlet_id,sku_id,reason,count(*) return_events,sum(return_quantity) returned_units,
 sum(return_amount) return_amount,sum(return_amount-returned_cogs) margin_reversal
FROM kpi_returns GROUP BY 1,2,3,4;

-- Business intent: prioritize billed debt, respecting historical terms and cutoff.
-- Grain: outlet/branch; pre-aggregated invoice snapshot avoids payment/SCD fanout.
-- DSO uses credit-only trailing-90-day billing, never the operating-period sales.
CREATE TEMP VIEW analysis_ar AS
WITH debt AS (
 SELECT branch_id,outlet_id,count(*) invoices,sum(outstanding) outstanding,
 sum(outstanding) FILTER(WHERE days_past_due>0) overdue,
 sum(outstanding) FILTER(WHERE days_past_due<=0) not_due,
 sum(outstanding) FILTER(WHERE days_past_due BETWEEN 1 AND 30) ar_1_30,
 sum(outstanding) FILTER(WHERE days_past_due BETWEEN 31 AND 60) ar_31_60,
 sum(outstanding) FILTER(WHERE days_past_due BETWEEN 61 AND 90) ar_61_90,
 sum(outstanding) FILTER(WHERE days_past_due>90) ar_over_90,
 sum(outstanding) FILTER(WHERE payment_method='Credit') credit_ar,
 sum(outstanding) FILTER(WHERE payment_method='Cash') cash_ar,sum(overpayment) overpayment,
 count(*) FILTER(WHERE outstanding>0 AND days_past_due>0) overdue_invoices,
 max(days_past_due) FILTER(WHERE outstanding>0) oldest_days_past_due
 FROM kpi_invoice_snapshot GROUP BY 1,2
), billings AS (
 SELECT branch_id,outlet_id,sum(invoice_amount) billings_90d FROM analytics.fact_invoices CROSS JOIN kpi_params p
 WHERE payment_method='Credit' AND invoice_date BETWEEN p.as_of_date-89 AND p.as_of_date GROUP BY 1,2
)
SELECT d.branch_id,d.outlet_id,d.invoices,d.outstanding,coalesce(d.overdue,0) overdue,
 coalesce(d.not_due,0) not_due,coalesce(ar_1_30,0) ar_1_30,coalesce(ar_31_60,0) ar_31_60,
 coalesce(ar_61_90,0) ar_61_90,coalesce(ar_over_90,0) ar_over_90,
 coalesce(credit_ar,0) credit_ar,coalesce(cash_ar,0) cash_ar,overpayment,
 overdue_invoices,oldest_days_past_due,coalesce(b.billings_90d,0) billings_90d,
 coalesce(d.overdue,0)/nullif(outstanding,0) overdue_rate,
 CASE WHEN p.dso_history_complete THEN coalesce(credit_ar,0)*90/nullif(b.billings_90d,0) END dso_90d,
 dense_rank() OVER(ORDER BY coalesce(d.overdue,0) DESC) collection_priority
FROM debt d LEFT JOIN billings b USING(branch_id,outlet_id) CROSS JOIN kpi_params p;

-- Business intent: compare observed productivity with recorded-sales mismatch.
-- Grain: recorded salesman; calls deduplicated before ratios; no causal claim.
CREATE TEMP VIEW analysis_salesman AS
WITH visits AS (
 SELECT salesman_id,count(*) scheduled,count(*) FILTER(WHERE is_completed) completed,
 count(*) FILTER(WHERE is_completed AND is_effective_call) effective,
 count(DISTINCT (salesman_id,visit_date,outlet_id)) FILTER(WHERE is_completed) completed_calls,
 count(DISTINCT (salesman_id,visit_date,outlet_id)) FILTER(WHERE is_completed AND has_attributed_order) converted_calls,
 count(DISTINCT (salesman_id,visit_date)) FILTER(WHERE is_completed) field_days
 FROM kpi_visits GROUP BY 1
), sales AS (
 SELECT salesman_id,count(*) sales_lines,count(*) FILTER(WHERE NOT is_salesman_match) mismatch_lines,
 sum(net_sales_before_returns) billed_sales,
 coalesce(sum(net_sales_before_returns) FILTER(WHERE NOT is_salesman_match),0) mismatch_sales
 FROM kpi_sales GROUP BY 1
), combined AS (
 SELECT d.salesman_id,d.salesman_name,d.home_branch_id,
 coalesce(v.scheduled,0) scheduled,coalesce(v.completed,0) completed,coalesce(v.effective,0) effective,
 coalesce(v.completed_calls,0) completed_calls,coalesce(v.converted_calls,0) converted_calls,coalesce(v.field_days,0) field_days,
 coalesce(s.sales_lines,0) sales_lines,coalesce(s.mismatch_lines,0) mismatch_lines,
 coalesce(s.billed_sales,0) billed_sales,coalesce(s.mismatch_sales,0) mismatch_sales
 FROM analytics.dim_salesman d LEFT JOIN visits v USING(salesman_id) LEFT JOIN sales s USING(salesman_id)
)
SELECT *,completed::numeric/nullif(scheduled,0) completion_rate,
 effective::numeric/nullif(completed,0) effective_call_rate,
 converted_calls::numeric/nullif(completed_calls,0) conversion_rate,
 completed::numeric/nullif(field_days,0) calls_per_day,
 mismatch_lines::numeric/nullif(sales_lines,0) mismatch_rate FROM combined;

-- Business intent: actionable, mutually exclusive snapshot outlet segments.
-- Rule precedence: inactive, champions, at risk, recent developing, other active.
-- Grain: outlet. These labels are descriptive; F/M scores remain relative.
CREATE TEMP VIEW analysis_rfm AS
SELECT r.*,o.branch_id,o.outlet_name,p.rfm_history_complete,
 CASE WHEN frequency=0 THEN 'Inactive'
 WHEN r_score>=4 AND f_score>=4 AND m_score>=4 THEN 'Champions'
 WHEN r_score<=2 THEN 'At risk'
 WHEN r_score>=4 THEN 'Recent developing' ELSE 'Other active' END segment
FROM kpi_rfm_scored r JOIN analytics.dim_outlet o USING(outlet_id) CROSS JOIN kpi_params p;
-- Cohort/month grain retained; window compares adjacent observed cells only.
CREATE TEMP VIEW analysis_cohort AS
SELECT *,active_outlets-lag(active_outlets) OVER(PARTITION BY cohort_month ORDER BY month_age) active_outlet_change
FROM kpi_cohort;
