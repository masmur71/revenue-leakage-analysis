-- JSON EXPLAIN harness; warmed, serial, fixed settings; three measured repetitions.
-- DDL for temporary views precedes read-only execution; no source mutation.
\set ON_ERROR_STOP on
\ir 00_candidates.sql
SET jit=off;
SET max_parallel_workers_per_gather=0;
SET work_mem='16MB';
SET statement_timeout='180s';
CREATE TEMP TABLE bench_evidence(payload jsonb);
BEGIN READ ONLY;
DO $$
DECLARE q text; v text; n int; plan json; mismatch bigint;
BEGIN
 FOREACH q IN ARRAY ARRAY['leakage','visits','cohort'] LOOP
  EXECUTE format('SELECT count(*) FROM ((TABLE bench_%1$s_before EXCEPT ALL TABLE bench_%1$s_after) UNION ALL (TABLE bench_%1$s_after EXCEPT ALL TABLE bench_%1$s_before)) d',q) INTO mismatch;
  IF mismatch<>0 THEN RAISE EXCEPTION 'Unequal results for %: % rows',q,mismatch; END IF;
  INSERT INTO bench_evidence VALUES(jsonb_build_object('kind','equality','query',q,'mismatched_rows',mismatch));
  -- Warm both candidates once; do not include warmup in timing summaries.
  FOREACH v IN ARRAY ARRAY['before','after'] LOOP
   EXECUTE format('EXPLAIN (ANALYZE, BUFFERS, FORMAT JSON) SELECT * FROM bench_%s_%s',q,v) INTO plan;
  END LOOP;
  FOR n IN 1..3 LOOP
   -- Reverse ordering in even repetitions to reduce systematic first-run bias.
   FOREACH v IN ARRAY CASE WHEN n%2=1 THEN ARRAY['before','after'] ELSE ARRAY['after','before'] END LOOP
    EXECUTE format('EXPLAIN (ANALYZE, BUFFERS, FORMAT JSON) SELECT * FROM bench_%s_%s',q,v) INTO plan;
    INSERT INTO bench_evidence VALUES(jsonb_build_object('kind','plan','query',q,'variant',v,'repetition',n,'explain',plan));
   END LOOP;
  END LOOP;
 END LOOP;
END $$;
SELECT payload::text FROM bench_evidence;
ROLLBACK;
