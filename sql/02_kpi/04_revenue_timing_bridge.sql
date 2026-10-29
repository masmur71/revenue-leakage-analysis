-- Purpose: reconcile event-period net sales to stored sale-cohort restatements.
-- Question: why do period amounts differ from fact_sales.net_sales?
-- Grain: enterprise reporting period. Source: sales lines and return events.
-- This diagnostic deliberately reads later returns to explain the stored
-- restatement; never use that restatement as a historical as-of KPI.
\ir 00_metric_context.sql
BEGIN TRANSACTION READ ONLY;
WITH components AS (
 SELECT c.net_sales AS event_net_sales,
 (SELECT coalesce(sum(net_sales),0) FROM kpi_sales) AS restated_net_sales,
 (SELECT coalesce(sum(r.return_amount),0) FROM kpi_returns r
  JOIN analytics.fact_sales s USING(sales_line_id) CROSS JOIN kpi_params p
  WHERE s.sales_date<p.period_start) AS prior_period_sales_returned_now,
 (SELECT coalesce(sum(r.return_amount),0) FROM analytics.fact_returns r
  JOIN kpi_sales s USING(sales_line_id) CROSS JOIN kpi_params p
  WHERE r.return_date>p.period_end) AS period_sales_returned_later
 FROM kpi_commercial c
)
SELECT *,event_net_sales+prior_period_sales_returned_now-period_sales_returned_later
 AS reconciled_restated_net_sales,
 restated_net_sales=event_net_sales+prior_period_sales_returned_now-period_sales_returned_later AS reconciles
FROM components;
ROLLBACK;
