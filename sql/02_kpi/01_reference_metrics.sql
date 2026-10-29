-- Purpose: produce reproducible KPI reference values, RFM inputs and cohort cells.
-- Question: what does each agreed KPI measure for this period/snapshot?
-- Sources/grain/assumptions: docs/kpi_dictionary.md and 00_metric_context.sql.
-- SQL notes: pre-aggregated child facts prevent invoice fan-out; numerator and
-- denominator support weighted rates rather than averages of percentages.
\ir 00_metric_context.sql
BEGIN TRANSACTION READ ONLY;
TABLE kpi_params;
SELECT * FROM kpi_metric_values ORDER BY metric_id;
SELECT count(*) AS eligible_outlets,count(*) FILTER(WHERE frequency>0) AS active_rfm_outlets,
 sum(frequency) AS outlet_order_count,sum(monetary) AS monetary FROM kpi_rfm;
SELECT r_score,f_score,m_score,count(*) outlets FROM kpi_rfm_scored GROUP BY 1,2,3 ORDER BY 1,2,3;
SELECT * FROM kpi_cohort ORDER BY cohort_month,month_age;
ROLLBACK;
