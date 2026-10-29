# Revenue Leakage & Sales Performance Analysis

SQL and Excel portfolio project using synthetic Indonesian FMCG distributor data. No real company data or business claims are included. The dataset covers 2024–2025; the original generator is unavailable, so reproduction uses the supplied PostgreSQL snapshot.

## Requirements and reproduction

Use Windows PowerShell, Docker Desktop running Linux containers, and PostgreSQL 16 client tools inside the container. Excel Windows with Power Query and Data Model is required in Phase 4. Never store passwords in this repository.

The current local container is `bintel-pg`, source database/user `bintel`. The source is read-only for analysis; `leakage_lab` is the isolated restored copy for later optimization. PostgreSQL CLI authentication here uses the existing container configuration.

To reproduce on a fresh machine with a local disposable PostgreSQL container (choose your own local authentication configuration), copy and restore the bundled dump. The destination database must not already exist:

```powershell
docker cp data/analytics.dump bintel-pg:/tmp/leakage-analytics.dump
docker exec bintel-pg createdb -U bintel -T template0 leakage_lab
docker exec bintel-pg pg_restore -U bintel -d leakage_lab --no-owner --no-privileges --exit-on-error /tmp/leakage-analytics.dump
docker cp sql/01_audit/01_audit.sql bintel-pg:/tmp/leakage-audit.sql
docker exec bintel-pg psql -U bintel -d leakage_lab -X -v ON_ERROR_STOP=1 -f /tmp/leakage-audit.sql
if ($LASTEXITCODE -ne 0) { throw 'Audit failed' }
```

For a fresh Docker installation, an example local-only container is:

```powershell
docker run --name bintel-pg -d -e POSTGRES_USER=bintel -e POSTGRES_DB=bintel -e POSTGRES_HOST_AUTH_METHOD=trust postgres:16
docker exec bintel-pg pg_isready -U bintel -d bintel
```

Wait until `pg_isready` reports accepting connections before restoring. This example publishes no ports and uses trust authentication only inside the disposable local container; it is not a deployment configuration. Use your existing container when available.

To create a new snapshot and copy from an existing source, run `./sql/00_setup/backup_restore.ps1`. It refuses to overwrite an existing copy database. The script exports only `analytics`, omitting ownership and privileges. No separate seed is required: `data/analytics.dump` contains schema and synthetic rows. DDL alone is in `sql/00_setup/analytics_schema.sql`.

## Current deliverables

- `docs/data_audit.md`: observed grain, relationships, controls and limitations.
- `docs/audit-source.txt` and `docs/audit-restored.txt`: identical audit evidence.
- `docs/schema-inventory.txt`: existing indexes, constraints and date coverage.
- `sql/01_audit/01_audit.sql`: executable audit with hard failure on invalid controls.
- `PROJECT_BRIEF.md`: agreed phases and exit criteria.

## KPI contract and verification

Phase 1 definitions are in `docs/kpi_dictionary.md`: 46 scalar metrics plus outlet RFM and onboarding-cohort retention. Default operating period is calendar 2025, with a 2025-12-31 snapshot. Run all contract checks and the source/copy comparison from PowerShell:

```powershell
./sql/02_kpi/03_verify_phase1.ps1
```

The runner checks five date scenarios, compares reference outputs against `leakage_lab`, and confirms invalid date parameters are rejected. It writes evidence to `docs/kpi-validation-*.txt` and `docs/kpi-reference-*.txt`. Both databases must exist (see restore instructions above).

For a different supported period, after the runner has copied scripts into the container:

```powershell
docker exec bintel-pg psql -U bintel -d bintel -X -v ON_ERROR_STOP=1 -v period_start=2024-01-01 -v period_end=2024-12-31 -v as_of_date=2024-12-31 -f /tmp/leakage-kpi/01_reference_metrics.sql
```

No final business recommendations or Excel dashboard have been produced yet. Phase 2 provides diagnostic analysis using the agreed contract. Keep monetary values numeric and distinguish transaction dates from snapshot dates. Period revenue recognizes returns on their return dates; AR does not assume merchandise returns are invoice-settling credit notes. See the dictionary before comparing figures from the earlier database summary.

## Core SQL diagnostics

Phase 2 analysis and limitations are documented in `docs/core_analysis.md`. Run:

```powershell
./sql/03_analysis/03_verify_phase2.ps1
```

The runner checks all diagnostic grains and financial reconciliations on five date scenarios, compares eight full-row fingerprints on source/copy, and produces `docs/analysis-report-2025.txt`. Query definitions are in `sql/03_analysis/00_analysis_context.sql`; `01_business_report.sql` answers the business questions and `02_validate_analysis.sql` enforces the controls. All report reads use a read-only transaction; only session-local metadata is created beforehand. Revenue leakage indicators remain separate; missing target volume/price components are explicitly disclosed. Phase 3 will benchmark selected queries on the isolated copy.

## Excel model and dashboard

Phase 4 delivers [RevenueLeakage.xlsx](excel/RevenueLeakage.xlsx), with Power Query, Data Model, explicit measures and interactive dashboards. Native numeric/refresh captures reconcile; final visual and manual recovery/reopen checks are operator-reported. See the [acceptance and handoff record](docs/excel/phase4-handoff.md) for evidence and attribution. Export with `./sql/05_excel/02_export.ps1`; typed M, DAX and relationships are in `excel/`. Follow the [Excel model guide](docs/excel_model.md) and [short Indonesian completion guide](docs/panduan_penyelesaian_phase4.md). Do not restart the exhaustive Verify suite or rebuild the final manually formatted dashboard just to check it.

## Query optimization evidence

Phase 3 benchmarks three diagnostic queries on the isolated copy. Run `./sql/04_optimization/03_run_benchmarks.ps1` to reproduce rewrite-only, composite-index and unhelpful-index comparisons. The runner enforces `leakage_lab`, records 54 measured JSON plans plus exact result comparisons over five date cases, and removes its experiment indexes. See `docs/optimization_report.md` for methodology, timing ranges, plan/buffer interpretation and tradeoffs; raw evidence is under `docs/optimization/`. Original reference queries remain unchanged.
