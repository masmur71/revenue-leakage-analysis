-- Session-local definitions precede the strictly read-only source transaction.
\set ON_ERROR_STOP on
\ir ../03_analysis/00_analysis_context.sql
-- One independently dated commercial event; no lifetime return columns.
CREATE TEMP VIEW excel_events AS
SELECT 'S:'||sales_line_id event_id,'Sale' event_type,sales_date event_date,
 branch_id,outlet_id,sku_id,salesman_id,channel,sales_order_id,
 quantity sales_units,0 returned_units,quantity*reference_price reference_sales,
 gross_sales,discount_amount,net_sales_before_returns billed_sales,cogs_sold,
 greatest(quantity*reference_price-gross_sales,0) price_shortfall,
 greatest(gross_sales-quantity*reference_price,0) price_upside,
 CASE WHEN promotion_id IS NOT NULL AND NOT is_promo_valid THEN discount_amount ELSE 0 END invalid_promo_discount,
 (promotion_id IS NOT NULL)::int promo_lines,
 (promotion_id IS NOT NULL AND NOT is_promo_valid)::int invalid_promo_lines,
 (NOT is_salesman_match)::int mismatch_lines,
 CASE WHEN NOT is_salesman_match THEN net_sales_before_returns ELSE 0 END mismatch_sales,
 0::numeric return_amount,0::numeric returned_cogs,NULL::text return_reason
FROM analytics.fact_sales CROSS JOIN kpi_params p WHERE sales_date<=p.as_of_date
UNION ALL
SELECT 'R:'||r.return_id,'Return',r.return_date,r.branch_id,r.outlet_id,r.sku_id,
 s.salesman_id,s.channel,NULL::text,0,r.return_quantity,0,0,0,0,0,0,0,0,0,0,0,0,
 r.return_amount,r.returned_cogs,r.reason
FROM analytics.fact_returns r JOIN analytics.fact_sales s USING(sales_line_id)
CROSS JOIN kpi_params p WHERE r.return_date<=p.as_of_date;

CREATE TEMP VIEW excel_visits AS
WITH sale_calls AS (SELECT DISTINCT salesman_id,sales_date,outlet_id FROM analytics.fact_sales)
SELECT v.visit_id,v.visit_date,v.branch_id,v.outlet_id,v.salesman_id,o.channel,
 v.is_completed,v.is_effective_call,c.outlet_id IS NOT NULL has_attributed_order,
 v.salesman_id||'|'||v.visit_date::text||'|'||v.outlet_id call_key,
 v.salesman_id||'|'||v.visit_date::text salesman_day_key
FROM analytics.fact_visits v JOIN analytics.dim_outlet o USING(outlet_id)
LEFT JOIN sale_calls c ON c.salesman_id=v.salesman_id AND c.sales_date=v.visit_date AND c.outlet_id=v.outlet_id
CROSS JOIN kpi_params p WHERE v.visit_date<=p.as_of_date;

-- Outlet-month rows preserve conformed geographical filters and fixed cohort denominators.
CREATE TEMP VIEW excel_cohort AS
WITH eligible AS (
 SELECT outlet_id,branch_id,channel,date_trunc('month',onboard_date)::date cohort_month
 FROM analytics.dim_outlet CROSS JOIN kpi_params p
 WHERE onboard_date BETWEEN DATE '2024-01-01' AND p.as_of_date
), activity AS (SELECT DISTINCT outlet_id,date_trunc('month',sales_date)::date activity_month FROM analytics.fact_sales)
SELECT e.outlet_id||'|'||m::date::text cohort_cell_id,e.outlet_id,e.branch_id,e.channel,e.cohort_month,
 m::date activity_month,(extract(year FROM age(m,e.cohort_month))*12+extract(month FROM age(m,e.cohort_month)))::int month_age,
 (a.outlet_id IS NOT NULL)::int active_outlet
FROM eligible e CROSS JOIN kpi_params p
CROSS JOIN LATERAL generate_series(e.cohort_month,date_trunc('month',p.as_of_date),interval '1 month') m
LEFT JOIN activity a ON a.outlet_id=e.outlet_id AND a.activity_month=m::date
WHERE (m+interval '1 month'-interval '1 day')::date<=p.as_of_date;

\ir 03_support_controls.sql
BEGIN ISOLATION LEVEL REPEATABLE READ READ ONLY;
SET LOCAL datestyle = 'ISO, YMD';
SET LOCAL statement_timeout = '180s';
\copy (SELECT * FROM analytics.dim_branch ORDER BY branch_id) TO '/tmp/leakage-phase4/export/DimBranch.csv' CSV HEADER ENCODING 'UTF8'
\copy (SELECT outlet_id,outlet_name,channel,onboard_date FROM analytics.dim_outlet ORDER BY outlet_id) TO '/tmp/leakage-phase4/export/DimOutlet.csv' CSV HEADER ENCODING 'UTF8'
\copy (SELECT sku_id,sku_name,category FROM analytics.dim_product ORDER BY sku_id) TO '/tmp/leakage-phase4/export/DimProduct.csv' CSV HEADER ENCODING 'UTF8'
\copy (SELECT salesman_id,salesman_name,home_branch_id FROM analytics.dim_salesman ORDER BY salesman_id) TO '/tmp/leakage-phase4/export/DimSalesman.csv' CSV HEADER ENCODING 'UTF8'
\copy (SELECT DISTINCT channel FROM analytics.dim_outlet ORDER BY channel) TO '/tmp/leakage-phase4/export/DimChannel.csv' CSV HEADER ENCODING 'UTF8'
\copy (SELECT date,year,quarter,month,month_name,to_char(date,'YYYY-MM') year_month FROM analytics.dim_date CROSS JOIN kpi_params p WHERE date BETWEEN DATE '2024-01-01' AND p.as_of_date ORDER BY date) TO '/tmp/leakage-phase4/export/DimDate.csv' CSV HEADER ENCODING 'UTF8'
\copy (SELECT * FROM excel_events ORDER BY event_id) TO '/tmp/leakage-phase4/export/FactCommercial.csv' CSV HEADER ENCODING 'UTF8'
\copy (SELECT * FROM excel_visits ORDER BY visit_id) TO '/tmp/leakage-phase4/export/FactVisits.csv' CSV HEADER ENCODING 'UTF8'
\copy (SELECT invoice_id,invoice_date,branch_id,outlet_id,salesman_id,channel,payment_method,invoice_amount,allocated_paid,outstanding,overpayment,applicable_term_days,contractual_due_date,days_past_due,CASE WHEN payment_method='Credit' AND invoice_date BETWEEN p.as_of_date-89 AND p.as_of_date THEN invoice_amount ELSE 0 END credit_billings_90d FROM kpi_invoice_snapshot CROSS JOIN kpi_params p ORDER BY invoice_id) TO '/tmp/leakage-phase4/export/FactInvoiceSnapshot.csv' CSV HEADER ENCODING 'UTF8'
\copy (SELECT t.target_id,t.branch_id,to_date(target_month||'-01','YYYY-MM-DD') month_start,target_net_sales FROM analytics.fact_targets t CROSS JOIN kpi_params p WHERE to_date(target_month||'-01','YYYY-MM-DD')<=p.as_of_date ORDER BY target_id) TO '/tmp/leakage-phase4/export/FactTargets.csv' CSV HEADER ENCODING 'UTF8'
\copy (SELECT pay.payment_id,pay.payment_date,i.branch_id,i.outlet_id,i.salesman_id,i.channel,pay.payment_amount,pay.allocation_status FROM analytics.fact_payments pay JOIN analytics.fact_invoices i USING(invoice_id) CROSS JOIN kpi_params p WHERE pay.payment_date<=p.as_of_date ORDER BY payment_id) TO '/tmp/leakage-phase4/export/FactPayments.csv' CSV HEADER ENCODING 'UTF8'
\copy (SELECT r.outlet_id,r.branch_id,o.channel,r.recency_days,r.frequency,r.monetary,r.r_score,r.f_score,r.m_score,r.segment,r.rfm_history_complete FROM analysis_rfm r JOIN analytics.dim_outlet o USING(outlet_id) ORDER BY outlet_id) TO '/tmp/leakage-phase4/export/FactRFM.csv' CSV HEADER ENCODING 'UTF8'
\copy (SELECT * FROM excel_cohort ORDER BY cohort_cell_id) TO '/tmp/leakage-phase4/export/FactCohort.csv' CSV HEADER ENCODING 'UTF8'
\copy (SELECT period_start,period_end,as_of_date,observed_rfm_start,rfm_history_complete,dso_history_complete FROM kpi_params) TO '/tmp/leakage-phase4/export/ReportContext.csv' CSV HEADER ENCODING 'UTF8'
\copy (SELECT * FROM kpi_metric_values ORDER BY metric_id) TO '/tmp/leakage-phase4/export/SQLReference.csv' CSV HEADER ENCODING 'UTF8'
\copy (SELECT * FROM analysis_branch_month ORDER BY branch_id,month) TO '/tmp/leakage-phase4/export/SQLBranchMonth.csv' CSV HEADER ENCODING 'UTF8'
\copy (SELECT * FROM analysis_yoy ORDER BY branch_id,month) TO '/tmp/leakage-phase4/export/SQLYoY.csv' CSV HEADER ENCODING 'UTF8'
\copy (SELECT * FROM analysis_cohort ORDER BY cohort_month,activity_month) TO '/tmp/leakage-phase4/export/SQLCohort.csv' CSV HEADER ENCODING 'UTF8'
\ir 01_filter_controls.sql
\copy (SELECT * FROM excel_support_controls ORDER BY case_id) TO '/tmp/leakage-phase4/export/SQLSupportControls.csv' CSV HEADER ENCODING 'UTF8'
ROLLBACK;
