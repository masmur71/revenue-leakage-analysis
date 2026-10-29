-- Purpose: audit source grain, relationships, SCD2 and financial controls.
-- Run with psql -X -v ON_ERROR_STOP=1. Source tables are read-only.
-- Historical valid_to is inclusive; NULL denotes an open interval.
CREATE TEMP TABLE audit_checks (check_name text, failures bigint);
BEGIN TRANSACTION READ ONLY;
DO $$
DECLARE r record; n bigint;
BEGIN
  FOR r IN SELECT c.conname, c.conrelid::regclass AS child,
      c.confrelid::regclass AS parent,
      (SELECT string_agg(format('s.%I = p.%I',a.attname,b.attname),' AND ')
       FROM unnest(c.conkey,c.confkey) AS k(src,dst)
       JOIN pg_attribute a ON a.attrelid=c.conrelid AND a.attnum=k.src
       JOIN pg_attribute b ON b.attrelid=c.confrelid AND b.attnum=k.dst) AS condition,
      (SELECT string_agg(format('s.%I IS NOT NULL',a.attname),' AND ')
       FROM unnest(c.conkey) k JOIN pg_attribute a ON a.attrelid=c.conrelid AND a.attnum=k) AS present
    FROM pg_constraint c WHERE c.connamespace='analytics'::regnamespace AND c.contype='f'
  LOOP
    EXECUTE format('SELECT count(*) FROM %s s WHERE %s AND NOT EXISTS (SELECT 1 FROM %s p WHERE %s)',r.child,r.present,r.parent,r.condition) INTO n;
    INSERT INTO audit_checks VALUES ('orphan: '||r.conname,n);
  END LOOP;
  FOR r IN SELECT tablename FROM pg_tables WHERE schemaname='analytics' ORDER BY tablename LOOP
    EXECUTE format('SELECT count(*) FROM analytics.%I',r.tablename) INTO n;
    RAISE NOTICE 'row_count | % | %',r.tablename,n;
  END LOOP;
END $$;
INSERT INTO audit_checks
SELECT 'sales financial identities',count(*) FROM analytics.fact_sales
WHERE gross_sales <> quantity * unit_price_before_discount
 OR net_sales_before_returns <> gross_sales-discount_amount
 OR net_sales <> net_sales_before_returns-return_amount
 OR net_cogs <> cogs_sold-returned_cogs OR gross_profit <> net_sales-net_cogs
UNION ALL
SELECT 'invoice vs sales lines',count(*) FROM analytics.fact_invoices i
LEFT JOIN (SELECT invoice_id,sum(net_sales_before_returns) amount FROM analytics.fact_sales GROUP BY 1) s USING(invoice_id)
WHERE i.invoice_amount IS DISTINCT FROM s.amount
UNION ALL
SELECT 'returns vs sales lines',count(*) FROM analytics.fact_sales s
LEFT JOIN (SELECT sales_line_id,sum(return_amount) amount,sum(return_quantity) qty,sum(returned_cogs) cost FROM analytics.fact_returns GROUP BY 1) r USING(sales_line_id)
WHERE s.return_amount IS DISTINCT FROM coalesce(r.amount,0)
 OR s.return_quantity IS DISTINCT FROM coalesce(r.qty,0)
 OR s.returned_cogs IS DISTINCT FROM coalesce(r.cost,0)
UNION ALL
SELECT 'credit term point-in-time coverage',count(*) FROM analytics.fact_invoices i
WHERE (SELECT count(*) FROM analytics.dim_outlet_credit_term d WHERE d.outlet_id=i.outlet_id AND i.invoice_date>=d.valid_from AND (d.valid_to IS NULL OR i.invoice_date<=d.valid_to))<>1
UNION ALL
SELECT 'assignment point-in-time coverage',count(*) FROM analytics.fact_sales s
WHERE (SELECT count(*) FROM analytics.dim_salesman_assignment d WHERE d.outlet_id=s.outlet_id AND s.sales_date>=d.valid_from AND (d.valid_to IS NULL OR s.sales_date<=d.valid_to))<>1
UNION ALL
SELECT 'credit SCD overlap',count(*) FROM analytics.dim_outlet_credit_term a JOIN analytics.dim_outlet_credit_term b ON a.outlet_id=b.outlet_id AND a.credit_term_sk<b.credit_term_sk
AND daterange(a.valid_from,a.valid_to,'[]') && daterange(b.valid_from,b.valid_to,'[]')
UNION ALL
SELECT 'assignment SCD overlap',count(*) FROM analytics.dim_salesman_assignment a JOIN analytics.dim_salesman_assignment b ON a.outlet_id=b.outlet_id AND a.assignment_sk<b.assignment_sk
AND daterange(a.valid_from,a.valid_to,'[]') && daterange(b.valid_from,b.valid_to,'[]')
UNION ALL
SELECT 'target branch-month duplicates',count(*) FROM (SELECT branch_id,target_month FROM analytics.fact_targets GROUP BY 1,2 HAVING count(*)>1) d
UNION ALL
SELECT 'effective call without sale',count(*) FROM analytics.fact_visits v WHERE is_effective_call AND NOT EXISTS (SELECT 1 FROM analytics.fact_sales s WHERE s.outlet_id=v.outlet_id AND s.sales_date=v.visit_date);
TABLE audit_checks;
SELECT min(sales_date) AS first_sale,max(sales_date) AS last_sale,sum(gross_sales) AS gross_sales,sum(discount_amount) AS discounts,sum(return_amount) AS returns,sum(net_sales) AS net_sales,sum(gross_profit) AS gross_profit FROM analytics.fact_sales;
SELECT min(payment_date),max(payment_date),allocation_status,count(*),sum(payment_amount) FROM analytics.fact_payments GROUP BY allocation_status;
SELECT table_name,column_name,data_type,is_nullable FROM information_schema.columns WHERE table_schema='analytics' ORDER BY table_name,ordinal_position;
DO $$ BEGIN IF EXISTS (SELECT 1 FROM audit_checks WHERE failures>0) THEN RAISE EXCEPTION 'Audit failed: inspect nonzero failures'; END IF; END $$;
ROLLBACK;
