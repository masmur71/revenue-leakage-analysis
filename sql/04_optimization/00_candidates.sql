-- Optimization candidates for the original synthetic dataset, not generated load.
-- Safety: run only on the isolated copy. Temporary view definitions only.
\set ON_ERROR_STOP on
DO $$ BEGIN
 IF current_database()<>'leakage_lab' THEN
  RAISE EXCEPTION 'Optimization requires the isolated leakage_lab database';
 END IF;
END $$;
\ir ../03_analysis/00_analysis_context.sql
CREATE TEMP VIEW bench_leakage_before AS SELECT * FROM analysis_leakage;
-- Preserve all metrics/windows; replace three copies of the event CTE with
-- GROUPING SETS. Aggregate once at each required grain, without a wide CTE spool.
CREATE TEMP VIEW bench_leakage_after AS
WITH totals AS (
 SELECT CASE WHEN grouping(branch_id)=0 THEN 'branch' WHEN grouping(sku_id)=0 THEN 'sku' ELSE 'outlet' END dimension,
 coalesce(CASE WHEN grouping(branch_id)=0 THEN branch_id END,
 CASE WHEN grouping(sku_id)=0 THEN sku_id END,CASE WHEN grouping(outlet_id)=0 THEN outlet_id END) member_id,
 sum(units) units,sum(reference_sales) reference_sales,sum(gross_sales) gross_sales,
 sum(price_shortfall) price_shortfall,sum(price_upside) price_upside,sum(discounts) discounts,
 sum(billed_sales) billed_sales,sum(cogs_sold) cogs_sold,sum(invalid_promo_discount) invalid_promo_discount,
 sum(returns) returns,sum(returned_cogs) returned_cogs,sum(returned_units) returned_units
 FROM analysis_events GROUP BY GROUPING SETS ((branch_id),(sku_id),(outlet_id))
)
SELECT *,billed_sales-returns net_sales,billed_sales-returns-cogs_sold+returned_cogs gross_profit,
 (billed_sales-returns-cogs_sold+returned_cogs)/nullif(billed_sales-returns,0) margin_rate,
 discounts/nullif(gross_sales,0) discount_rate,returns/nullif(billed_sales,0) return_rate,
 returns-returned_cogs return_margin_reversal,
 dense_rank() OVER(PARTITION BY dimension ORDER BY price_shortfall DESC) price_exposure_rank,
 sum(price_shortfall) OVER(PARTITION BY dimension ORDER BY price_shortfall DESC,member_id ROWS UNBOUNDED PRECEDING)
 /nullif(sum(price_shortfall) OVER(PARTITION BY dimension),0) cumulative_price_share FROM totals;

CREATE TEMP VIEW bench_visits_before AS SELECT * FROM kpi_visit_totals;
-- EXISTS in aggregate FILTER can repeatedly probe sales. Build unique order-call
-- keys once; LEFT JOIN retains cancelled/unconverted visits and event-grain counts.
CREATE TEMP VIEW bench_visits_after AS
WITH order_calls AS MATERIALIZED (
 SELECT DISTINCT salesman_id,sales_date,outlet_id FROM kpi_sales
), visits AS (
 SELECT v.*,s.outlet_id IS NOT NULL has_attributed_order FROM analytics.fact_visits v
 CROSS JOIN kpi_params p LEFT JOIN order_calls s ON s.salesman_id=v.salesman_id
 AND s.sales_date=v.visit_date AND s.outlet_id=v.outlet_id
 WHERE v.visit_date BETWEEN p.period_start AND p.period_end
)
SELECT count(*) scheduled,count(*) FILTER(WHERE is_completed) completed,
 count(*) FILTER(WHERE is_completed AND is_effective_call) effective,
 count(DISTINCT (salesman_id,visit_date,outlet_id)) FILTER(WHERE is_completed) completed_calls,
 count(DISTINCT (salesman_id,visit_date,outlet_id)) FILTER(WHERE is_completed AND has_attributed_order) converted_calls,
 count(DISTINCT (salesman_id,visit_date)) FILTER(WHERE is_completed) salesman_days FROM visits;

CREATE TEMP VIEW bench_cohort_before AS SELECT * FROM analysis_cohort;
-- Aggregate activity and cohort sizes before expanding the calendar. The grid is
-- cohort-month, not outlet-month; inactive outlets remain in the denominator.
CREATE TEMP VIEW bench_cohort_after AS
WITH eligible AS MATERIALIZED (
 SELECT outlet_id,date_trunc('month',onboard_date)::date cohort_month FROM analytics.dim_outlet CROSS JOIN kpi_params p
 WHERE onboard_date BETWEEN DATE '2024-01-01' AND p.as_of_date
), sizes AS (
 SELECT cohort_month,count(*) cohort_outlets FROM eligible GROUP BY 1
), activity AS (
 SELECT DISTINCT e.cohort_month,s.outlet_id,date_trunc('month',s.sales_date)::date activity_month
 FROM eligible e JOIN analytics.fact_sales s USING(outlet_id) CROSS JOIN kpi_params p
 WHERE s.sales_date<=p.as_of_date
), active AS (
 SELECT cohort_month,activity_month,count(*) active_outlets FROM activity GROUP BY 1,2
), cells AS (
 SELECT c.cohort_month,m::date activity_month,
 (extract(year FROM age(m,c.cohort_month))*12+extract(month FROM age(m,c.cohort_month)))::int month_age,
 c.cohort_outlets,coalesce(a.active_outlets,0) active_outlets,
 coalesce(a.active_outlets,0)::numeric/nullif(c.cohort_outlets,0) retention_rate
 FROM sizes c CROSS JOIN kpi_params p
 CROSS JOIN LATERAL generate_series(c.cohort_month,date_trunc('month',p.as_of_date),interval '1 month') m
 LEFT JOIN active a ON a.cohort_month=c.cohort_month AND a.activity_month=m::date
 WHERE (m+interval '1 month'-interval '1 day')::date<=p.as_of_date
)
SELECT *,active_outlets-lag(active_outlets) OVER(PARTITION BY cohort_month ORDER BY month_age) active_outlet_change FROM cells;
