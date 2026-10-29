-- Validate diagnostics against the independently defined Phase 1 contract.
-- Grain, additive components, date boundaries and exact YoY bridge must hold.
\set ON_ERROR_STOP on
\ir 00_analysis_context.sql
CREATE TEMP TABLE analysis_checks(check_name text,failures bigint);
BEGIN READ ONLY;
INSERT INTO analysis_checks
SELECT 'leakage grain unique',count(*) FROM (SELECT dimension,member_id FROM analysis_leakage GROUP BY 1,2 HAVING count(*)<>1) x
UNION ALL SELECT 'leakage additive controls per dimension',count(*) FROM (
 SELECT dimension,sum(reference_sales) ref,sum(price_shortfall) shortfall,sum(price_upside) upside,sum(discounts) discount,
 sum(invalid_promo_discount) invalid,sum(returns) returned,sum(net_sales) net,sum(gross_profit) profit,sum(units) units FROM analysis_leakage GROUP BY 1
) x CROSS JOIN kpi_commercial c WHERE (x.ref,x.shortfall,x.upside,x.discount,x.invalid,x.returned,x.net,x.profit,x.units)
 IS DISTINCT FROM (c.reference_sales,c.price_shortfall,c.price_upside,c.discounts,c.invalid_promo_discount,c.returns,c.net_sales,c.gross_profit,c.units)
UNION ALL SELECT 'member financial bridges',count(*) FROM analysis_leakage
 WHERE reference_sales+price_upside-price_shortfall-discounts<>billed_sales OR net_sales<>billed_sales-returns
 OR gross_profit<>net_sales-cogs_sold+returned_cogs
UNION ALL SELECT 'branch month unique',count(*) FROM (SELECT branch_id,month FROM analysis_branch_month GROUP BY 1,2 HAVING count(*)<>1) x
UNION ALL SELECT 'branch month sales profit units reconcile',count(*) FROM
 (SELECT sum(net_sales) net,sum(gross_profit) profit,sum(units) units FROM analysis_branch_month) x CROSS JOIN kpi_commercial c
 WHERE (x.net,x.profit,x.units) IS DISTINCT FROM (c.net_sales,c.gross_profit,c.units)
UNION ALL SELECT 'complete month target coverage',count(*) FROM analysis_branch_month WHERE complete_month AND target IS NULL
UNION ALL SELECT 'partial month targets undefined',count(*) FROM analysis_branch_month WHERE NOT complete_month AND target IS NOT NULL
UNION ALL SELECT 'YoY bridge exact within numeric division tolerance',count(*) FROM analysis_yoy
 WHERE volume_at_prior_blended_price IS NOT NULL AND blended_price_mix_change IS NOT NULL
 AND abs(net_change-volume_at_prior_blended_price-blended_price_mix_change-discount_change-return_change)>0.000001
UNION ALL SELECT 'return reasons preserve amount and events',count(*) FROM
 (SELECT coalesce(sum(return_amount),0) amount,coalesce(sum(return_events),0) events FROM analysis_return_reasons) x
 WHERE x.amount<>(SELECT returns FROM kpi_return_totals) OR x.events<>(SELECT count(*) FROM kpi_returns)
UNION ALL SELECT 'AR grain unique',count(*) FROM (SELECT branch_id,outlet_id FROM analysis_ar GROUP BY 1,2 HAVING count(*)<>1) x
UNION ALL SELECT 'AR additive controls',count(*) FROM
 (SELECT coalesce(sum(outstanding),0) outstanding,coalesce(sum(overdue),0) overdue,coalesce(sum(credit_ar),0) credit_ar,
 coalesce(sum(cash_ar),0) cash_ar,coalesce(sum(overpayment),0) overpayment,coalesce(sum(billings_90d),0) billings FROM analysis_ar) x
 CROSS JOIN kpi_ar a CROSS JOIN kpi_dso_denominator d
 WHERE (x.outstanding,x.overdue,x.credit_ar,x.cash_ar,x.overpayment,x.billings)
 IS DISTINCT FROM (a.outstanding,a.overdue,a.credit_outstanding,a.cash_outstanding,a.overpayment,d.credit_billings_90d)
UNION ALL SELECT 'AR aging partitions',count(*) FROM analysis_ar WHERE outstanding<>not_due+ar_1_30+ar_31_60+ar_61_90+ar_over_90 OR overdue<>ar_1_30+ar_31_60+ar_61_90+ar_over_90
UNION ALL SELECT 'salesman additive controls',count(*) FROM
 (SELECT sum(scheduled) scheduled,sum(completed) completed,sum(effective) effective,sum(completed_calls) calls,
 sum(converted_calls) converted,sum(field_days) days,sum(sales_lines) lines,sum(mismatch_lines) mismatch,sum(billed_sales) billed,sum(mismatch_sales) mismatch_sales FROM analysis_salesman) x
 CROSS JOIN kpi_visit_totals v CROSS JOIN kpi_sales_totals s
 WHERE (x.scheduled,x.completed,x.effective,x.calls,x.converted,x.days,x.lines,x.mismatch,x.billed,x.mismatch_sales)
 IS DISTINCT FROM (v.scheduled,v.completed,v.effective,v.completed_calls,v.converted_calls,v.salesman_days,s.sales_lines,s.mismatch_lines,s.billed_sales,s.mismatch_sales)
UNION ALL SELECT 'RFM segments preserve all outlet scores',count(*) FROM
 ((SELECT outlet_id,frequency,monetary,r_score,f_score,m_score FROM analysis_rfm EXCEPT ALL SELECT outlet_id,frequency,monetary,r_score,f_score,m_score FROM kpi_rfm_scored)
 UNION ALL (SELECT outlet_id,frequency,monetary,r_score,f_score,m_score FROM kpi_rfm_scored EXCEPT ALL SELECT outlet_id,frequency,monetary,r_score,f_score,m_score FROM analysis_rfm)) x
UNION ALL SELECT 'cohort cells preserved',count(*) FROM
 ((SELECT cohort_month,activity_month,month_age,cohort_outlets,active_outlets,retention_rate FROM analysis_cohort EXCEPT ALL SELECT * FROM kpi_cohort)
 UNION ALL (SELECT * FROM kpi_cohort EXCEPT ALL SELECT cohort_month,activity_month,month_age,cohort_outlets,active_outlets,retention_rate FROM analysis_cohort)) x;
SELECT * FROM analysis_checks ORDER BY check_name;
DO $$ BEGIN IF EXISTS(SELECT 1 FROM analysis_checks WHERE failures<>0) THEN RAISE EXCEPTION 'Phase 2 validation failed'; END IF; END $$;
-- Deterministic full-row fingerprints compare every diagnostic row on source/copy.
SELECT name,count(*) rows,md5(string_agg(payload,E'\n' ORDER BY payload)) fingerprint FROM (
 SELECT 'leakage' name,row_to_json(t)::text payload FROM analysis_leakage t
 UNION ALL SELECT 'branch_month',row_to_json(t)::text FROM analysis_branch_month t
 UNION ALL SELECT 'yoy',row_to_json(t)::text FROM analysis_yoy t
 UNION ALL SELECT 'returns',row_to_json(t)::text FROM analysis_return_reasons t
 UNION ALL SELECT 'ar',row_to_json(t)::text FROM analysis_ar t
 UNION ALL SELECT 'salesman',row_to_json(t)::text FROM analysis_salesman t
 UNION ALL SELECT 'rfm',row_to_json(t)::text FROM analysis_rfm t
 UNION ALL SELECT 'cohort',row_to_json(t)::text FROM analysis_cohort t
) all_rows GROUP BY name ORDER BY name;
ROLLBACK;
