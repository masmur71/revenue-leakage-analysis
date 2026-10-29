-- Purpose: validate KPI contracts against live data and explicit edge fixtures.
-- Run after Phase 0 audit; source tables are never modified.
-- The final exception makes a failing contract return a nonzero psql exit code.
\ir 00_metric_context.sql
CREATE TEMP TABLE metric_checks(check_name text PRIMARY KEY,failures bigint NOT NULL);
BEGIN TRANSACTION READ ONLY;
INSERT INTO metric_checks
SELECT 'invoice snapshot has one row per eligible invoice',abs(
 (SELECT count(*) FROM kpi_invoice_snapshot)-
 (SELECT count(*) FROM analytics.fact_invoices CROSS JOIN kpi_params p WHERE invoice_date<=p.as_of_date))
UNION ALL
SELECT 'historical term coverage',count(*) FROM analytics.fact_invoices i CROSS JOIN kpi_params p
WHERE invoice_date<=p.as_of_date AND (SELECT count(*) FROM analytics.dim_outlet_credit_term d
 WHERE d.outlet_id=i.outlet_id AND i.invoice_date>=d.valid_from AND (d.valid_to IS NULL OR i.invoice_date<=d.valid_to))<>1
UNION ALL
SELECT 'cash exception and credit due dates',count(*) FROM kpi_invoice_snapshot
WHERE applicable_term_days IS NULL OR credit_term_days IS DISTINCT FROM applicable_term_days
 OR due_date IS DISTINCT FROM contractual_due_date
UNION ALL
SELECT 'recognized payment domains',count(*) FROM analytics.fact_invoices WHERE payment_method NOT IN ('Cash','Credit')
UNION ALL
SELECT 'recognized allocation domain',count(*) FROM analytics.fact_payments WHERE allocation_status IS DISTINCT FROM 'Allocated'
UNION ALL
SELECT 'nonnegative invoice and payment amounts',
 (SELECT count(*) FROM analytics.fact_invoices WHERE invoice_amount<0)+
 (SELECT count(*) FROM analytics.fact_payments WHERE payment_amount<0)
UNION ALL
SELECT 'invoice snapshot balance identity',count(*) FROM kpi_invoice_snapshot
WHERE invoice_amount-allocated_paid<>outstanding-overpayment
UNION ALL
SELECT 'aging buckets partition receivables',count(*) FROM kpi_ar
WHERE outstanding<>not_due+overdue_1_30+overdue_31_60+overdue_61_90+overdue_90_plus
 OR overdue<>overdue_1_30+overdue_31_60+overdue_61_90+overdue_90_plus
 OR outstanding<>credit_outstanding+cash_outstanding
UNION ALL
SELECT 'financial and reference price bridges',count(*) FROM kpi_commercial
WHERE gross_sales-discounts<>billed_sales OR billed_sales-returns<>net_sales
 OR net_sales-net_cogs<>gross_profit
 OR reference_sales+price_upside-price_shortfall-discounts<>billed_sales
UNION ALL
SELECT 'return sources preserve original sales dimensions',count(*) FROM analytics.fact_returns r
JOIN analytics.fact_sales s USING(sales_line_id)
WHERE r.outlet_id<>s.outlet_id OR r.branch_id<>s.branch_id OR r.sku_id<>s.sku_id OR r.return_date<s.sales_date
UNION ALL
SELECT 'positive quantity and reference price',count(*) FROM analytics.fact_sales WHERE quantity<=0 OR reference_price<=0
UNION ALL
SELECT 'sales amount domains',count(*) FROM analytics.fact_sales WHERE gross_sales<0 OR discount_amount<0 OR discount_amount>gross_sales
UNION ALL
SELECT 'invalid promo discount is a subset',count(*) FROM kpi_commercial
WHERE invalid_promo_discount<0 OR invalid_promo_discount>discounts OR invalid_promo_lines>promo_lines
UNION ALL
SELECT 'salesman match flag agrees with IDs',count(*) FROM analytics.fact_sales
WHERE is_salesman_match IS DISTINCT FROM (salesman_id=assigned_salesman_id)
UNION ALL
SELECT 'effective calls require completion and attributed sale',count(*) FROM kpi_visits
WHERE is_effective_call AND (NOT is_completed OR NOT has_attributed_order)
UNION ALL
SELECT 'conversion and completion counts bounded',count(*) FROM kpi_visit_totals
WHERE effective>completed OR completed>scheduled OR converted_calls>completed_calls
UNION ALL
SELECT 'onboarding precedes first transaction',count(*) FROM analytics.fact_sales s
JOIN analytics.dim_outlet o USING(outlet_id) WHERE s.sales_date<o.onboard_date
UNION ALL
SELECT 'target format and nonnegative amount',count(*) FROM analytics.fact_targets
WHERE target_month!~'^[0-9]{4}-(0[1-9]|1[0-2])$' OR target_net_sales<0
UNION ALL
SELECT 'target month matches date dimension',count(*) FROM analytics.fact_targets t
JOIN analytics.dim_date d USING(date_sk) WHERE to_char(d.date,'YYYY-MM')<>t.target_month
UNION ALL
SELECT 'target coverage is complete for all branch months',count(*) FROM analytics.dim_branch b CROSS JOIN kpi_params p
CROSS JOIN LATERAL generate_series(date_trunc('month',p.period_start),date_trunc('month',p.period_end),interval '1 month') m
WHERE (SELECT count(*) FROM analytics.fact_targets t WHERE t.branch_id=b.branch_id AND t.target_month=to_char(m,'YYYY-MM'))<>1
UNION ALL
SELECT 'unique metric IDs',count(*) FROM (SELECT metric_id FROM kpi_metric_values GROUP BY 1 HAVING count(*)<>1) d
UNION ALL
SELECT 'zero denominators produce NULL',count(*) FROM kpi_metric_values WHERE denominator=0 AND value IS NOT NULL
UNION ALL
SELECT 'partial-month targets are undefined',count(*) FROM kpi_metric_values CROSS JOIN kpi_params p
WHERE metric_id IN ('target_net_sales','achievement_rate') AND value IS NOT NULL
 AND (p.period_start<>date_trunc('month',p.period_start)::date
 OR p.period_end<>(date_trunc('month',p.period_end)+interval '1 month'-interval '1 day')::date)
UNION ALL
SELECT 'incomplete DSO history is undefined',count(*) FROM kpi_metric_values CROSS JOIN kpi_params p
WHERE metric_id='dso_90d' AND NOT p.dso_history_complete AND value IS NOT NULL
UNION ALL
SELECT 'RFM frequency reconciles to source orders',abs(
 (SELECT sum(frequency) FROM kpi_rfm)-
 (SELECT count(DISTINCT (outlet_id,sales_order_id)) FROM analytics.fact_sales CROSS JOIN kpi_params p WHERE sales_date BETWEEN p.as_of_date-364 AND p.as_of_date))::bigint
UNION ALL
SELECT 'RFM monetary reconciles to event net sales',CASE WHEN
 (SELECT sum(monetary) FROM kpi_rfm)=
 (SELECT coalesce(sum(net_sales_before_returns),0) FROM analytics.fact_sales CROSS JOIN kpi_params p WHERE sales_date BETWEEN p.as_of_date-364 AND p.as_of_date)-
 (SELECT coalesce(sum(return_amount),0) FROM analytics.fact_returns CROSS JOIN kpi_params p WHERE return_date BETWEEN p.as_of_date-364 AND p.as_of_date)
 THEN 0 ELSE 1 END
UNION ALL
SELECT 'cohort numerator bounded and no future cells',count(*) FROM kpi_cohort CROSS JOIN kpi_params p
WHERE active_outlets>cohort_outlets OR retention_rate NOT BETWEEN 0 AND 1
 OR (activity_month+interval '1 month'-interval '1 day')::date>p.as_of_date
UNION ALL
SELECT 'cohort denominator stable across months',count(*) FROM
 (SELECT cohort_month FROM kpi_cohort GROUP BY 1 HAVING min(cohort_outlets)<>max(cohort_outlets)) d
UNION ALL
SELECT 'RFM scores respect activity and bounds',count(*) FROM kpi_rfm_scored
WHERE (frequency=0 AND (r_score<>0 OR f_score<>0 OR m_score<>0))
 OR (frequency>0 AND (r_score NOT BETWEEN 1 AND 5 OR f_score NOT BETWEEN 1 AND 5 OR m_score NOT BETWEEN 1 AND 5))
UNION ALL
SELECT 'RFM frequency ties share scores',count(*) FROM
 (SELECT frequency FROM kpi_rfm_scored GROUP BY frequency HAVING min(f_score)<>max(f_score)) d
UNION ALL
SELECT 'RFM monetary ties share active scores',count(*) FROM
 (SELECT monetary FROM kpi_rfm_scored WHERE frequency>0 GROUP BY monetary HAVING min(m_score)<>max(m_score)) d;

-- Edge fixture: invoice issued 100, payment 40 at cutoff, another 60 later,
-- merchandise return 10. Expected AR is 60, not 0 or 50.
WITH payments(payment_date,amount,status) AS (VALUES
 (DATE '2025-01-31',40::numeric,'Allocated'),
 (DATE '2025-02-01',60::numeric,'Allocated'),
 (DATE '2025-01-30',5::numeric,'Unallocated'))
INSERT INTO metric_checks SELECT 'fixture: payments respect snapshot and allocation',
 CASE WHEN 100-coalesce(sum(amount) FILTER(WHERE payment_date<=DATE '2025-01-31' AND status='Allocated'),0)=60 THEN 0 ELSE 1 END FROM payments;
WITH boundaries(days_past_due,expected) AS (VALUES (-1,'not_due'),(0,'not_due'),(1,'1_30'),(30,'1_30'),(31,'31_60'),(60,'31_60'),(61,'61_90'),(90,'61_90'),(91,'over_90'))
INSERT INTO metric_checks SELECT 'fixture: aging boundary inclusivity',count(*) FROM boundaries
WHERE CASE WHEN days_past_due<=0 THEN 'not_due' WHEN days_past_due<=30 THEN '1_30' WHEN days_past_due<=60 THEN '31_60' WHEN days_past_due<=90 THEN '61_90' ELSE 'over_90' END<>expected;
WITH versions(valid_from,valid_to) AS (VALUES (DATE '2024-01-01',DATE '2024-12-31'),(DATE '2025-01-01',DATE '2025-12-31')),
 events(event_date) AS (VALUES (DATE '2024-12-31'),(DATE '2025-01-01'))
INSERT INTO metric_checks SELECT 'fixture: inclusive SCD2 year boundary',count(*) FROM events
WHERE (SELECT count(*) FROM versions WHERE event_date BETWEEN valid_from AND valid_to)<>1;
WITH events(kind,event_date,amount) AS (VALUES ('sale',DATE '2024-12-31',100::numeric),('return',DATE '2025-01-02',10::numeric))
INSERT INTO metric_checks SELECT 'fixture: later return does not rewrite earlier sales',
 CASE WHEN sum(CASE WHEN kind='sale' THEN amount ELSE -amount END) FILTER(WHERE event_date<=DATE '2024-12-31')=100
 AND sum(CASE WHEN kind='sale' THEN amount ELSE -amount END) FILTER(WHERE event_date BETWEEN DATE '2025-01-01' AND DATE '2025-01-31')=-10 THEN 0 ELSE 1 END FROM events;
WITH promo(promotion_id,is_promo_valid,discount) AS (VALUES (NULL::text,false,20::numeric),('valid',true,10::numeric),('invalid',false,5::numeric))
INSERT INTO metric_checks SELECT 'fixture: non-promo false flag excluded',CASE WHEN sum(discount) FILTER(WHERE promotion_id IS NOT NULL AND is_promo_valid=false)=5 THEN 0 ELSE 1 END FROM promo;
INSERT INTO metric_checks VALUES ('fixture: zero denominator is undefined',CASE WHEN 1::numeric/nullif(0,0) IS NULL THEN 0 ELSE 1 END);
-- This fixed reconciliation belongs only to the entire observed history.
INSERT INTO metric_checks SELECT 'full-history Phase 0 controls',count(*) FROM kpi_commercial CROSS JOIN kpi_params p
WHERE p.period_start=DATE '2024-01-01' AND p.period_end=DATE '2025-12-31'
 AND (gross_sales<>28960053341.78 OR net_sales<>28537677278.72 OR gross_profit<>6493439139.72);
SELECT * FROM metric_checks ORDER BY check_name;
DO $$ BEGIN IF EXISTS (SELECT 1 FROM metric_checks WHERE failures<>0) THEN RAISE EXCEPTION 'KPI validation failed'; END IF; END $$;
ROLLBACK;
