-- Shared KPI contract. Include from 01_reference_metrics.sql or 02_validate_metrics.sql.
-- Sources: analytics facts/dimensions; grain and assumptions: docs/kpi_dictionary.md.
-- Only session-local views are created. No permanent schema or source data is changed.
\set ON_ERROR_STOP on
\if :{?period_start}
\else
\set period_start '2025-01-01'
\endif
\if :{?period_end}
\else
\set period_end '2025-12-31'
\endif
\if :{?as_of_date}
\else
\set as_of_date '2025-12-31'
\endif

CREATE TEMP VIEW kpi_params AS
SELECT :'period_start'::date AS period_start, :'period_end'::date AS period_end,
       :'as_of_date'::date AS as_of_date,
       greatest(:'as_of_date'::date-364,DATE '2024-01-01') AS observed_rfm_start,
       :'as_of_date'::date-364>=DATE '2024-01-01' AS rfm_history_complete,
       :'as_of_date'::date-89>=DATE '2024-01-01' AS dso_history_complete;
DO $$ BEGIN
 IF EXISTS (SELECT 1 FROM kpi_params WHERE period_start>period_end OR period_end>as_of_date
  OR period_start<DATE '2024-01-01' OR as_of_date>DATE '2025-12-31')
 THEN RAISE EXCEPTION 'Require 2024-01-01 <= period_start <= period_end <= as_of_date <= 2025-12-31'; END IF;
END $$;

CREATE TEMP VIEW kpi_sales AS
SELECT s.* FROM analytics.fact_sales s CROSS JOIN kpi_params p
WHERE s.sales_date BETWEEN p.period_start AND p.period_end;
CREATE TEMP VIEW kpi_returns AS
SELECT r.*,s.invoice_id FROM analytics.fact_returns r
JOIN analytics.fact_sales s USING(sales_line_id) CROSS JOIN kpi_params p
WHERE r.return_date BETWEEN p.period_start AND p.period_end;

-- Each event is attributed to its own event date. Never reuse lifetime return
-- columns from fact_sales for historical or period-specific net sales.
CREATE TEMP VIEW kpi_sales_totals AS
SELECT coalesce(sum(gross_sales),0) gross_sales,
 coalesce(sum(discount_amount),0) discounts,
 coalesce(sum(net_sales_before_returns),0) billed_sales,
 coalesce(sum(cogs_sold),0) cogs_sold,
 coalesce(sum(quantity),0) units,
 coalesce(sum(quantity*reference_price),0) reference_sales,
 coalesce(sum(greatest(quantity*reference_price-gross_sales,0)),0) price_shortfall,
 coalesce(sum(greatest(gross_sales-quantity*reference_price,0)),0) price_upside,
 coalesce(sum(discount_amount) FILTER(WHERE promotion_id IS NOT NULL AND is_promo_valid=false),0) invalid_promo_discount,
 count(*) FILTER(WHERE promotion_id IS NOT NULL AND is_promo_valid=false) invalid_promo_lines,
 count(*) FILTER(WHERE promotion_id IS NOT NULL) promo_lines,
 coalesce(sum(net_sales_before_returns) FILTER(WHERE is_salesman_match=false),0) mismatch_sales,
 count(*) FILTER(WHERE is_salesman_match=false) mismatch_lines,
 count(*) sales_lines,count(DISTINCT outlet_id) active_outlets,
 count(DISTINCT sales_order_id) orders FROM kpi_sales;
CREATE TEMP VIEW kpi_return_totals AS
SELECT coalesce(sum(return_amount),0) returns,
 coalesce(sum(returned_cogs),0) returned_cogs,
 coalesce(sum(return_quantity),0) returned_units FROM kpi_returns;
CREATE TEMP VIEW kpi_commercial AS
SELECT s.*,r.*,billed_sales-returns AS net_sales,
 cogs_sold-returned_cogs AS net_cogs,
 billed_sales-returns-cogs_sold+returned_cogs AS gross_profit
FROM kpi_sales_totals s CROSS JOIN kpi_return_totals r;

-- Pre-aggregate payments at invoice grain to prevent fan-out. Only allocated
-- payments known at the snapshot count. Returns are not proven credit notes.
CREATE TEMP VIEW kpi_invoice_snapshot AS
WITH paid AS (
 SELECT invoice_id,sum(payment_amount) paid FROM analytics.fact_payments CROSS JOIN kpi_params p
 WHERE allocation_status='Allocated' AND payment_date<=p.as_of_date GROUP BY 1
)
SELECT i.*,coalesce(paid.paid,0) allocated_paid,
 greatest(i.invoice_amount-coalesce(paid.paid,0),0) outstanding,
 greatest(coalesce(paid.paid,0)-i.invoice_amount,0) overpayment,
 CASE WHEN i.payment_method='Cash' THEN 0 ELSE t.credit_term_days END AS applicable_term_days,
 i.invoice_date+CASE WHEN i.payment_method='Cash' THEN 0 ELSE t.credit_term_days END AS contractual_due_date,
 p.as_of_date-i.due_date AS days_past_due
FROM analytics.fact_invoices i CROSS JOIN kpi_params p LEFT JOIN paid USING(invoice_id)
LEFT JOIN analytics.dim_outlet_credit_term t ON t.outlet_id=i.outlet_id
 AND i.invoice_date>=t.valid_from AND (t.valid_to IS NULL OR i.invoice_date<=t.valid_to)
WHERE i.invoice_date<=p.as_of_date;
CREATE TEMP VIEW kpi_ar AS
SELECT coalesce(sum(outstanding),0) outstanding,
 coalesce(sum(outstanding) FILTER(WHERE days_past_due>0),0) overdue,
 coalesce(sum(outstanding) FILTER(WHERE days_past_due<=0),0) not_due,
 coalesce(sum(outstanding) FILTER(WHERE days_past_due BETWEEN 1 AND 30),0) overdue_1_30,
 coalesce(sum(outstanding) FILTER(WHERE days_past_due BETWEEN 31 AND 60),0) overdue_31_60,
 coalesce(sum(outstanding) FILTER(WHERE days_past_due BETWEEN 61 AND 90),0) overdue_61_90,
 coalesce(sum(outstanding) FILTER(WHERE days_past_due>90),0) overdue_90_plus,
 coalesce(sum(outstanding) FILTER(WHERE payment_method='Credit'),0) credit_outstanding,
 coalesce(sum(outstanding) FILTER(WHERE payment_method='Cash'),0) cash_outstanding,
 coalesce(sum(overpayment),0) overpayment FROM kpi_invoice_snapshot;
CREATE TEMP VIEW kpi_dso_denominator AS
SELECT coalesce(sum(invoice_amount),0) credit_billings_90d
FROM analytics.fact_invoices i CROSS JOIN kpi_params p
WHERE payment_method='Credit' AND invoice_date BETWEEN p.as_of_date-89 AND p.as_of_date;

CREATE TEMP VIEW kpi_visits AS
SELECT v.*,EXISTS (SELECT 1 FROM analytics.fact_sales s
 WHERE s.outlet_id=v.outlet_id AND s.salesman_id=v.salesman_id AND s.sales_date=v.visit_date) AS has_attributed_order
FROM analytics.fact_visits v CROSS JOIN kpi_params p
WHERE visit_date BETWEEN p.period_start AND p.period_end;
CREATE TEMP VIEW kpi_visit_totals AS
SELECT count(*) scheduled,count(*) FILTER(WHERE is_completed) completed,
 count(*) FILTER(WHERE is_completed AND is_effective_call) effective,
 count(DISTINCT (salesman_id,visit_date,outlet_id)) FILTER(WHERE is_completed) completed_calls,
 count(DISTINCT (salesman_id,visit_date,outlet_id)) FILTER(WHERE is_completed AND has_attributed_order) converted_calls,
 count(DISTINCT (salesman_id,visit_date)) FILTER(WHERE is_completed) salesman_days
FROM kpi_visits;
CREATE TEMP VIEW kpi_target AS
SELECT sum(target_net_sales) target FROM analytics.fact_targets t CROSS JOIN kpi_params p
WHERE to_date(t.target_month||'-01','YYYY-MM-DD') BETWEEN p.period_start AND p.period_end;
CREATE TEMP VIEW kpi_collections AS
SELECT coalesce(sum(payment_amount),0) allocated_receipts
FROM analytics.fact_payments CROSS JOIN kpi_params p
WHERE allocation_status='Allocated' AND payment_date BETWEEN p.period_start AND p.period_end;

-- RFM uses 365 calendar days ending at the snapshot. Restrict the customer
-- population to outlets onboarded by the snapshot, including zero-order outlets.
CREATE TEMP VIEW kpi_rfm AS
WITH orders AS (
 SELECT outlet_id,max(sales_date) last_order_date,count(DISTINCT sales_order_id) frequency,
 sum(net_sales_before_returns) billed FROM analytics.fact_sales CROSS JOIN kpi_params p
 WHERE sales_date BETWEEN p.as_of_date-364 AND p.as_of_date GROUP BY 1
), returns AS (
 SELECT outlet_id,sum(return_amount) returned FROM analytics.fact_returns CROSS JOIN kpi_params p
 WHERE return_date BETWEEN p.as_of_date-364 AND p.as_of_date GROUP BY 1
)
SELECT o.outlet_id,p.as_of_date-orders.last_order_date recency_days,
 coalesce(orders.frequency,0) frequency,coalesce(orders.billed,0)-coalesce(returns.returned,0) monetary
FROM analytics.dim_outlet o CROSS JOIN kpi_params p
LEFT JOIN orders USING(outlet_id) LEFT JOIN returns USING(outlet_id)
WHERE o.onboard_date<=p.as_of_date;
CREATE TEMP VIEW kpi_rfm_scored AS
WITH active AS (
 SELECT outlet_id,
 ceil(5*cume_dist() OVER(ORDER BY frequency))::integer f_score,
 ceil(5*cume_dist() OVER(ORDER BY monetary))::integer m_score
 FROM kpi_rfm WHERE frequency>0
)
SELECT r.*,
 CASE WHEN frequency=0 THEN 0 WHEN recency_days<=30 THEN 5 WHEN recency_days<=90 THEN 4
 WHEN recency_days<=180 THEN 3 WHEN recency_days<=270 THEN 2 ELSE 1 END AS r_score,
 coalesce(a.f_score,0) f_score,coalesce(a.m_score,0) m_score
FROM kpi_rfm r LEFT JOIN active a USING(outlet_id);

-- Complete calendar months only, including zero-activity cells. Pre-2024
-- onboarding cohorts are left-censored and excluded from this retention panel.
CREATE TEMP VIEW kpi_cohort AS
WITH eligible AS (
 SELECT o.outlet_id,date_trunc('month',onboard_date)::date cohort_month
 FROM analytics.dim_outlet o CROSS JOIN kpi_params p
 WHERE onboard_date BETWEEN DATE '2024-01-01' AND p.as_of_date
), activity AS (
 SELECT DISTINCT outlet_id,date_trunc('month',sales_date)::date activity_month
 FROM analytics.fact_sales CROSS JOIN kpi_params p WHERE sales_date<=p.as_of_date
), grid AS (
 SELECT e.*,m::date activity_month
 FROM eligible e CROSS JOIN kpi_params p
 CROSS JOIN LATERAL generate_series(e.cohort_month,date_trunc('month',p.as_of_date),interval '1 month') m
 WHERE (m+interval '1 month'-interval '1 day')::date<=p.as_of_date
)
SELECT g.cohort_month,g.activity_month,
 (extract(year FROM age(g.activity_month,g.cohort_month))*12+extract(month FROM age(g.activity_month,g.cohort_month)))::int month_age,
 count(*) cohort_outlets,count(a.outlet_id) active_outlets,
 count(a.outlet_id)::numeric/nullif(count(*),0) retention_rate
FROM grid g LEFT JOIN activity a ON a.outlet_id=g.outlet_id AND a.activity_month=g.activity_month
GROUP BY 1,2;

-- Long-form metric interface: ratios stored as fractions; currency as IDR.
-- Numerator/denominator remain visible for weighted aggregation and validation.
CREATE TEMP VIEW kpi_metric_values AS
SELECT v.metric_id,v.numerator,v.denominator,v.numerator/nullif(v.denominator,0) value,v.unit
FROM kpi_commercial c CROSS JOIN kpi_ar a CROSS JOIN kpi_dso_denominator d
CROSS JOIN kpi_visit_totals t CROSS JOIN kpi_target target CROSS JOIN kpi_params p
CROSS JOIN kpi_collections collections
CROSS JOIN LATERAL (VALUES
 ('gross_sales',c.gross_sales,1::numeric,'IDR'),
 ('discount_amount',c.discounts,1,'IDR'),
 ('billed_sales',c.billed_sales,1,'IDR'),
 ('return_amount',c.returns,1,'IDR'),
 ('net_sales',c.net_sales,1,'IDR'),
 ('net_cogs',c.net_cogs,1,'IDR'),
 ('gross_profit',c.gross_profit,1,'IDR'),
 ('gross_margin_rate',c.gross_profit,c.net_sales,'ratio'),
 ('discount_rate',c.discounts,c.gross_sales,'ratio'),
 ('return_rate',c.returns,c.billed_sales,'ratio'),
 ('sales_units',c.units,1,'units'),
 ('returned_units',c.returned_units,1,'units'),
 ('return_margin_reversal',c.returns-c.returned_cogs,1,'IDR'),
 ('reference_sales',c.reference_sales,1,'IDR'),
 ('price_shortfall',c.price_shortfall,1,'IDR'),
 ('price_upside',c.price_upside,1,'IDR'),
 ('price_deviation_rate',c.gross_sales-c.reference_sales,c.reference_sales,'ratio'),
 ('invalid_promo_discount',c.invalid_promo_discount,1,'IDR'),
 ('invalid_promo_line_rate',c.invalid_promo_lines,c.promo_lines,'ratio'),
 ('salesman_mismatch_sales',c.mismatch_sales,1,'IDR'),
 ('salesman_mismatch_rate',c.mismatch_lines,c.sales_lines,'ratio'),
 ('active_outlets',c.active_outlets,1,'outlets'),
 ('orders',c.orders,1,'orders'),
 ('average_order_value',c.billed_sales,c.orders,'IDR/order'),
 ('target_net_sales',CASE WHEN p.period_start=date_trunc('month',p.period_start)::date AND p.period_end=(date_trunc('month',p.period_end)+interval '1 month'-interval '1 day')::date THEN target.target END,1,'IDR'),
 ('achievement_rate',c.net_sales,CASE WHEN p.period_start=date_trunc('month',p.period_start)::date AND p.period_end=(date_trunc('month',p.period_end)+interval '1 month'-interval '1 day')::date THEN target.target END,'ratio'),
 ('outstanding_ar',a.outstanding,1,'IDR'),
 ('overdue_ar',a.overdue,1,'IDR'),
 ('overdue_rate',a.overdue,a.outstanding,'ratio'),
 ('ar_not_due',a.not_due,1,'IDR'),
 ('ar_1_30',a.overdue_1_30,1,'IDR'),
 ('ar_31_60',a.overdue_31_60,1,'IDR'),
 ('ar_61_90',a.overdue_61_90,1,'IDR'),
 ('ar_over_90',a.overdue_90_plus,1,'IDR'),
 ('credit_ar',a.credit_outstanding,1,'IDR'),
 ('cash_ar',a.cash_outstanding,1,'IDR'),
 ('overpayment',a.overpayment,1,'IDR'),
 ('allocated_receipts',collections.allocated_receipts,1,'IDR'),
 ('dso_90d',CASE WHEN p.as_of_date-89>=DATE '2024-01-01' THEN a.credit_outstanding*90 END,d.credit_billings_90d,'days'),
 ('scheduled_visits',t.scheduled,1,'visits'),
 ('completed_visits',t.completed,1,'visits'),
 ('effective_calls',t.effective,1,'visits'),
 ('visit_completion_rate',t.completed,t.scheduled,'ratio'),
 ('effective_call_rate',t.effective,t.completed,'ratio'),
 ('order_conversion_rate',t.converted_calls,t.completed_calls,'ratio'),
 ('calls_per_salesman_day',t.completed,t.salesman_days,'visits/day')
) v(metric_id,numerator,denominator,unit);
