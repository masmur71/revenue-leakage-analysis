-- Full multiset equivalence for five historical date scenarios (runner supplies dates).
\set ON_ERROR_STOP on
\ir 00_candidates.sql
SET jit=off;
SET max_parallel_workers_per_gather=0;
SET work_mem='16MB';
BEGIN READ ONLY;
DO $$ DECLARE q text; mismatch bigint; BEGIN
 FOREACH q IN ARRAY ARRAY['leakage','visits','cohort'] LOOP
  EXECUTE format('SELECT count(*) FROM ((TABLE bench_%1$s_before EXCEPT ALL TABLE bench_%1$s_after) UNION ALL (TABLE bench_%1$s_after EXCEPT ALL TABLE bench_%1$s_before)) d',q) INTO mismatch;
  IF mismatch<>0 THEN RAISE EXCEPTION 'Unequal results for %: % rows',q,mismatch; END IF;
  RAISE NOTICE 'PASS: % exact multiset equality',q;
 END LOOP;
END $$;
ROLLBACK;
