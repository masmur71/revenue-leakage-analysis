# Query optimization — Phase 3

## Scope and reproduction

The experiment uses the original synthetic snapshot in **leakage_lab only**: 329,263 sales rows, 105,918 visit rows, 6,725 returns and 3,000 outlets. No generated load rows were added. The source `bintel` is not queried or modified by the runner. Candidate SQL refuses to run outside `leakage_lab` before creating temporary views. The benchmark does not alter Phase 1 formulas or Phase 2 deliverables.

Run from PowerShell with the existing Docker container available:

```powershell
./sql/04_optimization/03_run_benchmarks.ps1
```

The runner requires the restored copy from README and its existing indexes. It refuses to proceed if either experiment index name already exists, refreshes planner statistics on four copy tables once, runs the three stages below, tests exact multiset equality over five date cases, and removes the two indexes it created in a `finally` block. No existing index is dropped. A failed run exits with an error; inspect the evidence file rather than treating partial output as a completed benchmark. After a forcibly interrupted session, inspect experiment indexes before rerunning; the refusal protects pre-existing objects. Temporary candidate views last only for a psql session. All timed query reads run in an explicit read-only transaction.

| Stage | Additional indexes | Purpose |
|---|---|---|
| baseline | None | Compare original and rewritten query on the original index set. |
| call_index | `(salesman_id, sales_date, outlet_id)` on sales | Test matching all three same-day attributed-call predicates in one index. Measure original and rewritten variants independently. |
| unhelpful_index | Call index plus `(is_promo_valid)` on sales | Measure a low-cardinality flag index that does not supply a selective predicate for these queries. Retain the call index to isolate the incremental change. |

The experiment index names are `leakage_opt_call_match` and `leakage_opt_promo_flag`. `index-sizes.txt` records their storage before cleanup. Adding an index can consume storage and impose maintenance costs; this read-only workload measures storage and read performance, not write overhead.

## Comparable conditions

PostgreSQL version and a timestamp are saved in `environment.json`. All stages set `jit=off`, `max_parallel_workers_per_gather=0`, `work_mem=16MB`, and a 180-second statement timeout per benchmark statement. These controlled settings make this experiment repeatable; they are not a production tuning recommendation. No planner scan method is disabled or forced.

Each query/variant receives one explicit warmup per stage, followed by three measured `EXPLAIN (ANALYZE, BUFFERS, FORMAT JSON)` executions. Original/rewritten order reverses on the even repetition. Equality checks precede the warmup and also warm data. This is a **warm-cache** experiment: no server restart, shared-buffer reset or operating-system cache flush occurs. Index construction can warm the cache further. Read results cannot support a cold-cache claim. Stages run sequentially, not randomly; other machine activity and stage order remain possible confounders. Medians and min/max ranges are reported, and plan/buffer evidence determines whether a plausible change accompanies elapsed-time variation. No significance claim is made from three repetitions.

`baseline.jsonl`, `call_index.jsonl` and `unhelpful_index.jsonl` retain every measured plan, including all scan/join/aggregate nodes, estimates, actual rows/loops and buffer/temp I/O. `timings.csv` extracts execution/planning time, root buffer counters and result rows for 54 measurements (3 queries x 2 variants x 3 repeats x 3 stages). `summary.csv` contains per-stage medians and min/max execution times. Root buffer counters already include descendants and must not be summed with child counters. Execution time includes EXPLAIN instrumentation and excludes client transfer/rendering and temporary-view DDL.

## Candidate logic and correctness

| Query | Original | Rewrite | Preserved semantics |
|---|---|---|---|
| Leakage | `analysis_leakage`: event CTE expanded into three dimension cuts, then aggregate/window | `GROUPING SETS` over the same event stream before the same windows | All branch/SKU/outlet amounts, margin/return ratios, dense ranks, running shares, return-event dates and invalid-attached-promo exclusions. |
| Visits | `kpi_visit_totals`: correlated sales `EXISTS` for attributed orders | Materialize unique salesman/date/outlet sale keys once and left join visit rows | Scheduled/completed/effective visit events, distinct completed/converted calls and distinct observed salesman-days. Cancelled/unconverted visits survive; no join fanout. |
| Cohort | `analysis_cohort`: outlet-by-month grid followed by aggregation | Aggregate cohort denominators and unique outlet/month activity first, then generate a cohort-by-month grid | Fixed onboarding denominator including never-buyers, complete observed months, zero activity cells, pre-2024 exclusion, month age and adjacent active-outlet change. |

`00_candidates.sql` contains executable before/after views; `01_benchmark.sql` captures plans and validates equality in every stage. `02_validate.sql` compares **both directions using EXCEPT ALL**, preserving duplicate multiplicity and NULL semantics, over calendar 2025, calendar 2024, mid-2025, partial January 2024 and the full observed period. This checks complete results, not rounded totals or sampled rows. Evidence is in `equality-*.txt`. The original KPI views remain the reference implementation; no formulas are silently replaced.

Materialized views are not introduced. The measured candidates operate on the existing dataset, and a persisted refresh/storage/staleness tradeoff has not been measured. A temporary materialized CTE used by a query is not a persisted materialized view.

## Measured results

Median execution time in milliseconds; range is min–max across three runs. All numbers are from the saved plans.

| Stage | Query | Original median (range) | Rewrite median (range) | Original / rewrite |
|---|---|---:|---:|---:|
| baseline | leakage | 2338.171 (2176.535–2395.875) | 709.465 (680.641–710.685) | 3.30x |
| baseline | visits | 1751.811 (1640.634–1788.372) | 583.98 (548.889–632.114) | 3.00x |
| baseline | cohort | 286.612 (267.413–420.77) | 78.563 (54.655–81.532) | 3.65x |
| call_index | leakage | 2150.135 (2149.553–2437.397) | 724.978 (658.912–728.646) | 2.97x |
| call_index | visits | 694.714 (603.847–731.261) | 339.786 (329.998–352.916) | 2.04x |
| call_index | cohort | 289.16 (272.779–290.562) | 139.694 (137.15–140.501) | 2.07x |
| unhelpful_index | leakage | 2517.963 (2154.213–2578.272) | 750.164 (747.087–824.744) | 3.36x |
| unhelpful_index | visits | 633.464 (581.032–666.467) | 335.455 (296.083–337.112) | 1.89x |
| unhelpful_index | cohort | 240.692 (230.963–243.008) | 137.493 (124.386–138.585) | 1.75x |

### Why the plans changed

Leakage: the rewrite reduces repeated wide-event expansion and repeated fact access. In baseline repetition 1, original shared buffers are 35,586 hits plus 26,520 reads, with 5,638 temporary blocks written; the rewrite uses 11,926 hits plus 8,776 reads and writes no temporary blocks. Both read/hit mixes reflect cache pressure despite warmups: warmed does not mean every page stays resident. With the call index stage, all pages in these runs are buffer hits, yet the rewrite still accesses 20,702 blocks versus 62,106 and avoids the original spill. The rewrite-only median improves by 3.30x. Block counts are accesses, not distinct physical pages or OS disk reads.

Visits: baseline repetition 1 executes correlated scans using `idx_fact_sales_outlet_id`; root shared hits are 393,058. The rewrite builds distinct sale-call keys and merge joins once, using 10,336 shared hits. With `leakage_opt_call_match`, the original uses an index-only lookup on all matching attributes (164,877 hits), while the rewrite uses an index-only scan with Unique plus merge join (2,056 hits). The rewrite-only median improves by 3.00x. The index additionally reduces the rewritten median from 583.980 to 339.786 ms (about 1.72x within this run), and the original from 1751.811 to 694.714 ms. Read-only index-only visibility is specific to this restored dataset; write-heavy workloads can require heap fetches.

Cohort: the cohort-level calendar avoids an outlet-level grid and retains fixed outlet denominators. Rewrite-only median improves by 3.65x. The added call index is also selected as a covering index for cohort sales, reducing buffers in repetition 1 to 970 in both variants versus 9,250 original / 4,904 rewrite without it. However, the rewritten cohort median **worsens** from 78.563 to 139.694 ms. Fewer buffers do not guarantee lower runtime; changed scan/join plans and system variability matter. The call index is a measured visit-workload option, not a universal speedup.

The deliberately unhelpful promo-flag index is absent from every measured plan. Node/index choices and root buffer/temp counters remain unchanged from the call-index stage for all query/variant pairs. Timings move in both directions despite this absence; the variation is not evidence that the promo index speeds these queries. It consumes 2,310,144 bytes (2,256 kB) with no observed access-path benefit. The call index consumes 7,774,208 bytes (7,592 kB). Both are removed at the end, so there is no permanent experimental index overhead.

Representative plan excerpts are in `docs/optimization/plan-excerpts.txt`; complete nested JSON plans remain the authoritative evidence. The report does not rely on elapsed time alone or cherry-pick a single repetition.

## Decision and exit verification

Keep the three rewritten candidates as tested alternatives for later export design; retain Phase 1/2 reference SQL unchanged for traceability. Prefer the leakage and cohort rewrites for these full-result scans. Consider the composite sales-call index only when the measured visit workload justifies its storage and the observed cohort regression is acceptable; it is not installed permanently here. Reject the promo-flag index for this workload. No materialized view or generated stress dataset is justified by this experiment.

All three stage checks compare every result row with exact multiset equality. Five additional date scenarios also pass for all three queries. The post-experiment Phase 0 audit passes all financial, relationship, SCD and row-count controls on the copy, and no experiment indexes remain. A wrong-database guard test on `postgres` fails with psql exit code 3 before candidate definitions, without touching `bintel`. Evidence is in `post-audit.txt`, `database-guard.txt`, and `equality-*.txt` under `docs/optimization/`.

Phase 3 exit criteria are fulfilled: three queries benchmarked with repeated before/after EXPLAIN ANALYZE BUFFERS evidence, rewrite and index effects separately measured, one unhelpful index demonstrated, and equal query results verified across the agreed historical scenarios. Limitations: one machine, small synthetic data, warmups without cache flushing, fixed serial/JIT-off settings, three repetitions, sequential index stages, no concurrent/writing workload, and no cold-cache or production-latency claim.
