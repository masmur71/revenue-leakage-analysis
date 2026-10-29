# SQL catalog

All paths are relative to the repository root. Original KPI/analysis SQL remains authoritative and unchanged. Scripts use `psql` include paths relative to their own file (`\ir`); use `-X -v ON_ERROR_STOP=1`. Parameters and reporting semantics are in the [KPI contract](kpi_dictionary.md). Source reads are read-only; temporary definitions precede read-only transactions because PostgreSQL forbids temporary-view DDL inside them.

| SQL file | Purpose / business question | Sources and assumptions |
|---|---|---|
| `sql/00_setup/analytics_schema.sql` | Owner-free source DDL; what objects are required? | Generated PostgreSQL 16 snapshot schema, no rows. Use only in an empty destination; full dump already includes DDL. |
| `sql/01_audit/01_audit.sql` | Are grain, FKs, SCD2 and money internally consistent? | All 13 analytics tables; 40 hard-failing controls, inclusive intervals; temporary check table only. |
| `sql/02_kpi/00_metric_context.sql` | Shared definitions; what does each KPI measure? | Sales/return/payment/visit/invoice facts and dated dimensions; event dates, allocated payments, invoice-date terms, fixed historical snapshot. Creates temporary views only. |
| `sql/02_kpi/01_reference_metrics.sql` | Produce 46 scalar KPI references, RFM and cohort controls. | Includes metric context; aggregate components before ratios, NULL for undefined; read-only report. |
| `sql/02_kpi/02_validate_metrics.sql` | Do definitions and edge fixtures reconcile? | Metric context plus session-local check table; no source mutations. Five supported date cases and explicit fixture boundaries. |
| `sql/02_kpi/04_revenue_timing_bridge.sql` | Why does event revenue differ from lifetime sale restatement? | Sales and later linked returns; intentionally reads later events for explanation, never for as-of revenue. |
| `sql/03_analysis/00_analysis_context.sql` | Define branch/region, leakage, YoY, returns, AR, salesman, RFM and cohort diagnostics. | Includes unchanged KPI views; event union avoids fanout; price includes product mix; target causes not identified. |
| `sql/03_analysis/01_business_report.sql` | Which members need investigation? | Analysis context; separate overlapping diagnostic cuts, money plus denominators, snapshot AR; read-only. |
| `sql/03_analysis/02_validate_analysis.sql` | Do diagnostic grains, bridges and eight full-result fingerprints reconcile? | Independent KPI controls; five date scenarios; hard failures and deterministic row fingerprints. |
| `sql/03_analysis/04_recommendation_evidence.sql` | Reproduce recommendation bases and scenario arithmetic. | Default 2025 analysis views; assumed 5% return prevention, 50% price correction, 5% overdue cash acceleration are scenarios, not observed benefits. |
| `sql/04_optimization/00_candidates.sql` | Can equivalent leakage/visit/cohort logic do less work? | Guarded to leakage_lab; temporary before/after views, GROUPING SETS and aggregate-first rewrites; no formula changes. |
| `sql/04_optimization/01_benchmark.sql` | Measure plans and buffers under comparable settings. | Candidate views; fixed warm-cache settings, one warmup and three measured repetitions per variant/stage; read-only timings. |
| `sql/04_optimization/02_validate.sql` | Are both variants exactly equivalent? | Bidirectional EXCEPT ALL preserves duplicates/NULL; runner supplies five historical cases. |
| `sql/05_excel/00_export.sql` | Export typed model facts/dimensions and independent controls. | Analytics facts/dimensions plus KPI/analysis views; repeatable-read read-only snapshot, full history through cutoff; COPY files only. Includes filter/support SQL below. |
| `sql/05_excel/01_filter_controls.sql` | Expected commercial results for selected filter combinations. | Independent fact aggregations, not exported facts; inclusive dates, NULL member filter means all; inherits export transaction. |
| `sql/05_excel/03_support_controls.sql` | Expected AR/visit/RFM filter results. | Reference views; fixed snapshot, stable RFM bands, consistent numerator/denominator filters; inherits export transaction. |

PowerShell entry points: backup_restore exports a source snapshot and refuses an existing copy; Phase 1/2 verification scripts compare source/copy; optimization runner owns temporary experimental indexes only; Excel exporter writes coherent CSV inputs. [Phase 6 reproduction](../sql/06_qa/01_reproduce.ps1) restores the bundled dump into a new QA database, runs the reference suites, and hashes independently reproduced exports. [Artifact verification](../sql/06_qa/02_verify_artifacts.mjs) rechecks native captures, workbook hash, documentation links and tracked-file hygiene without opening Excel.
