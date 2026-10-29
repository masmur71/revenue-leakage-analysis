# Setup

Run from the repository root in PowerShell. Use Docker Desktop with Linux containers, PostgreSQL 16, and Excel desktop for Windows with Power Query/Data Model support. Node.js 20.16+ is needed for validation.

## 1. Restore the snapshot

On a fresh machine, create a disposable local container:

```powershell
docker run --name bintel-pg -d -e POSTGRES_USER=bintel -e POSTGRES_DB=bintel -e POSTGRES_HOST_AUTH_METHOD=trust postgres:16
docker exec bintel-pg pg_isready -U bintel -d bintel
```

Wait for accepting connections. This local example publishes no host port and uses trust authentication; it is not a deployment configuration. Skip container creation if it already exists.

Restore only into the empty `bintel` database. Never restore over an existing source:

```powershell
docker cp data/analytics.dump bintel-pg:/tmp/leakage-analytics.dump
docker exec bintel-pg pg_restore -U bintel -d bintel --no-owner --no-privileges --exit-on-error /tmp/leakage-analytics.dump
if ($LASTEXITCODE -ne 0) { throw 'Restore failed' }
```

The dump includes schema and synthetic rows. Do not run the standalone DDL before restoring it. Existing source analysis remains read-only.

## 2. Prepare Excel

```powershell
./sql/05_excel/02_export.ps1
node excel/verify_exports.mjs
```

Require successful export and validation before opening the workbook. CSVs are generated locally in `data/excel/`; the tracked manifest records the accepted hashes. Follow the [refresh guide](refresh.md) to set the folder parameter and refresh.

## 3. Optional verification

```powershell
./sql/06_qa/01_reproduce.ps1
node sql/06_qa/02_verify_artifacts.mjs
```

The first command restores a new `leakage_qa_phase6` database and compares it with the source and accepted CSVs. It refuses an existing destination; reruns need a new `-Database leakage_qa_run2` name. Preserve the accepted manifest before comparison on an existing workspace. Detailed logs regenerate locally and are ignored by Git. No database is dropped or dashboard rebuilt. See [verification results](final_qa.md).

For query benchmarking, restore a separate empty `leakage_lab` from the same dump and run `./sql/04_optimization/03_run_benchmarks.ps1`. Methodology and retained plans are in the [optimization report](optimization_report.md).
