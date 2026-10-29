-- Additional native filter acceptance controls, evaluated from reference views.
CREATE TEMP VIEW excel_support_controls AS
WITH cases(case_id,branch_id,outlet_id,channel) AS (
 VALUES ('support_all',NULL::text,NULL::text,NULL::text),
 ('support_jakarta','B001',NULL,NULL),('support_bandung','B002',NULL,NULL),
 ('support_channel',NULL,NULL,(SELECT min(channel) FROM analytics.dim_outlet))
)
SELECT c.*,a.outstanding_ar,a.overdue_ar,a.credit_ar,a.credit_billings_90d,
 a.overdue_ar/nullif(a.outstanding_ar,0) overdue_rate,
 CASE WHEN p.dso_history_complete THEN a.credit_ar*90/nullif(a.credit_billings_90d,0) END dso_90d,
 v.scheduled_visits,v.completed_visits,v.effective_calls,
 v.completed_visits::numeric/nullif(v.scheduled_visits,0) visit_completion_rate,
 v.effective_calls::numeric/nullif(v.completed_visits,0) effective_call_rate,
 v.converted_calls::numeric/nullif(v.completed_calls,0) order_conversion_rate,
 v.completed_visits::numeric/nullif(v.salesman_days,0) calls_per_salesman_day,
 r.rfm_outlets,r.rfm_monetary
FROM cases c CROSS JOIN kpi_params p
CROSS JOIN LATERAL (
 SELECT coalesce(sum(outstanding),0) outstanding_ar,
 coalesce(sum(outstanding) FILTER(WHERE days_past_due>0),0) overdue_ar,
 coalesce(sum(outstanding) FILTER(WHERE payment_method='Credit'),0) credit_ar,
 coalesce(sum(invoice_amount) FILTER(WHERE payment_method='Credit' AND invoice_date BETWEEN p.as_of_date-89 AND p.as_of_date),0) credit_billings_90d
 FROM kpi_invoice_snapshot i
 WHERE (c.branch_id IS NULL OR i.branch_id=c.branch_id)
 AND (c.outlet_id IS NULL OR i.outlet_id=c.outlet_id) AND (c.channel IS NULL OR i.channel=c.channel)
) a
CROSS JOIN LATERAL (
 SELECT count(*) scheduled_visits,count(*) FILTER(WHERE is_completed) completed_visits,
 count(*) FILTER(WHERE is_completed AND is_effective_call) effective_calls,
 count(DISTINCT (salesman_id,visit_date,v.outlet_id)) FILTER(WHERE is_completed) completed_calls,
 count(DISTINCT (salesman_id,visit_date,v.outlet_id)) FILTER(WHERE is_completed AND has_attributed_order) converted_calls,
 count(DISTINCT (salesman_id,visit_date)) FILTER(WHERE is_completed) salesman_days
 FROM kpi_visits v JOIN analytics.dim_outlet o USING(outlet_id)
 WHERE (c.branch_id IS NULL OR v.branch_id=c.branch_id)
 AND (c.outlet_id IS NULL OR v.outlet_id=c.outlet_id) AND (c.channel IS NULL OR o.channel=c.channel)
) v
CROSS JOIN LATERAL (
 SELECT count(*) rfm_outlets,coalesce(sum(monetary),0) rfm_monetary
 FROM analysis_rfm r JOIN analytics.dim_outlet o USING(outlet_id)
 WHERE (c.branch_id IS NULL OR r.branch_id=c.branch_id)
 AND (c.outlet_id IS NULL OR r.outlet_id=c.outlet_id) AND (c.channel IS NULL OR o.channel=c.channel)
) r;
