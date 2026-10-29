-- Phase 2 business report. Intent/assumptions/grain/logic are documented in
-- 00_analysis_context.sql. Validation: 02_validate_analysis.sql and Phase 1.
\set ON_ERROR_STOP on
\ir 00_analysis_context.sql
BEGIN READ ONLY;
\echo LEAKAGE: enterprise bridge (IDR); indicators overlap and are not recoverable loss
SELECT * FROM kpi_commercial;
\echo LEAKAGE: ten highest price-shortfall members in each dimension
SELECT dimension,member_id,price_shortfall,discounts,invalid_promo_discount,returns,return_margin_reversal,net_sales,gross_profit,margin_rate,return_rate,cumulative_price_share
FROM analysis_leakage WHERE price_exposure_rank<=10 ORDER BY dimension,price_exposure_rank,member_id;
\echo LEAKAGE: largest discount exposures per dimension, with observed margin
SELECT dimension,member_id,discounts,discount_rate,net_sales,gross_profit,margin_rate FROM (
 SELECT *,row_number() OVER(PARTITION BY dimension ORDER BY discounts DESC,member_id) exposure_order FROM analysis_leakage
) ranked WHERE exposure_order<=10 ORDER BY dimension,exposure_order;
\echo TARGET: branch achievement; summed components, complete months only
SELECT branch_id,branch_name,sum(net_sales) net_sales,sum(target) target,sum(net_sales)-sum(target) gap,
 sum(net_sales)/nullif(sum(target),0) achievement_rate,count(*) FILTER(WHERE net_sales<target) below_target_months
FROM analysis_branch_month WHERE complete_month GROUP BY 1,2 ORDER BY gap;
\echo TARGET: region achievement
SELECT region_id,region_name,sum(net_sales) net_sales,sum(target) target,sum(net_sales)/nullif(sum(target),0) achievement_rate
FROM analysis_branch_month WHERE complete_month GROUP BY 1,2 ORDER BY achievement_rate;
\echo YEAR ON YEAR: descriptive volume and blended price including product mix, not target attribution
SELECT branch_id,sum(net_change) net_change,sum(volume_at_prior_blended_price) volume_component,
 sum(blended_price_mix_change) blended_price_mix_component,sum(discount_change) discount_component,sum(return_change) return_component
FROM analysis_yoy GROUP BY 1 ORDER BY net_change;
\echo RETURNS: reasons by event date
SELECT reason,sum(return_events) return_events,sum(return_amount) return_amount,sum(margin_reversal) margin_reversal FROM analysis_return_reasons GROUP BY 1 ORDER BY return_amount DESC;
\echo RETURNS: top outlet and SKU return rates; denominator and amount shown, no minimum-volume inference
SELECT dimension,member_id,billed_sales,returns,return_rate FROM (
 SELECT *,row_number() OVER(PARTITION BY dimension ORDER BY return_rate DESC NULLS LAST,member_id) return_order
 FROM analysis_leakage WHERE dimension IN ('outlet','sku') AND returns>0
) ranked WHERE return_order<=10 ORDER BY dimension,return_order;
\echo AR: branch balances, aging, credit-only DSO; no assumed credit notes
SELECT branch_id,sum(outstanding) outstanding,sum(overdue) overdue,sum(not_due) not_due,
 sum(ar_1_30) ar_1_30,sum(ar_31_60) ar_31_60,sum(ar_61_90) ar_61_90,sum(ar_over_90) ar_over_90,
 sum(cash_ar) cash_ar,sum(overpayment) overpayment,
 CASE WHEN bool_and(p.dso_history_complete) THEN sum(credit_ar)*90/nullif(sum(billings_90d),0) END dso_90d
FROM analysis_ar CROSS JOIN kpi_params p GROUP BY 1 ORDER BY overdue DESC;
\echo AR: highest overdue outlet balances
SELECT * FROM analysis_ar ORDER BY overdue DESC,outlet_id LIMIT 10;
\echo SALESMAN: observed productivity and mismatch, full population
SELECT * FROM analysis_salesman ORDER BY billed_sales DESC,salesman_id;
\echo RFM: mutually exclusive segments at snapshot
SELECT segment,rfm_history_complete,count(*) outlets,sum(frequency) outlet_orders,sum(monetary) monetary FROM analysis_rfm GROUP BY 1,2 ORDER BY 1;
\echo COHORT: weighted same-month activation and retention by observed age; age populations differ
SELECT month_age,count(*) observed_cohorts,sum(cohort_outlets) eligible_outlets,sum(active_outlets) active_outlets,
 sum(active_outlets)::numeric/nullif(sum(cohort_outlets),0) retention_rate FROM analysis_cohort GROUP BY 1 ORDER BY 1;
ROLLBACK;
