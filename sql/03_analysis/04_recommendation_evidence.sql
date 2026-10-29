-- Phase 5 focused evidence. Existing KPI definitions remain authoritative.
\set ON_ERROR_STOP on
\set period_start 2025-01-01
\set period_end 2025-12-31
\set as_of_date 2025-12-31
\ir ../03_analysis/00_analysis_context.sql
BEGIN READ ONLY;
\echo E1: observed commercial amounts; not a recoverable-loss total
SELECT net_sales,gross_profit,returns,returns-returned_cogs AS return_margin_reversal,
 discounts,price_shortfall,price_upside,invalid_promo_discount,mismatch_sales
FROM kpi_commercial;
\echo E2: return prevention base; national reasons overlap branch cuts
SELECT reason,sum(return_amount) AS return_amount,sum(margin_reversal) AS margin_reversal
FROM analysis_return_reasons GROUP BY reason ORDER BY return_amount DESC;
SELECT member_id,returns,billed_sales,return_rate FROM analysis_leakage
WHERE dimension='branch' AND member_id='B002';
\echo E3: price-control base and upper-bound scenario
SELECT member_id,price_shortfall,round(price_shortfall*0.5,2) AS assumed_revenue_ceiling
FROM analysis_leakage WHERE dimension='branch' AND price_shortfall>0;
\echo E4: cash opportunity; not incremental revenue
SELECT sum(outstanding) AS outstanding,sum(overdue) AS overdue,
 round(sum(overdue)*0.05,2) AS assumed_cash_acceleration FROM analysis_ar;
SELECT branch_id,sum(overdue) AS overdue FROM analysis_ar
GROUP BY branch_id ORDER BY overdue DESC LIMIT 1;
\echo E5: scenario arithmetic; revenue and GP are alternative lenses, never additive
SELECT round(returns*0.05,2) AS prevented_return_revenue,
 round((returns-returned_cogs)*0.05,2) AS retained_return_gp
FROM kpi_commercial;
ROLLBACK;
