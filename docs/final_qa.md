# Final QA and reproduction record — Phase 6

## Scope and execution

Reproduction starts from the committed synthetic `data/analytics.dump` in a newly created database, not from the existing restored copy. The [PowerShell runner](../sql/06_qa/01_reproduce.ps1) refuses an existing destination, restores into `template0`, and retains the QA database for inspection. It never drops a database or modifies source tables. The source `bintel` is audited and queried through unchanged reference SQL; temporary views/check tables are session-local. The accepted workbook is not rebuilt, opened or saved by this check.

Run from repository-root PowerShell after preparing the default CSVs as described in [README](../README.md):

```powershell
./sql/06_qa/01_reproduce.ps1
node sql/06_qa/02_verify_artifacts.mjs
git diff --check
```

The default destination is `leakage_qa_phase6`. For subsequent runs choose a new `-Database leakage_qa_run2` (or another unused `leakage_qa_*` name). With no accepted local CSVs, export them and run `node excel/verify_exports.mjs` first. Do not regenerate the accepted manifest immediately before a comparison on an existing workspace.

## Acceptance evidence

**Executed outcome: PASS.** On 30 October 2026 (Asia/Jakarta), `leakage_qa_phase6` restored successfully; all five source/copy scenarios and all 20 export hashes matched. Artifact verification passed for 561 saved native values, documentation links, the unchanged workbook and 54 optimization plans. The finalized runner also passed syntax validation and refused the existing QA database before restore; see [safety-guard.txt](qa/safety-guard.txt).

Machine-readable outcomes are in [reproduction.json](qa/reproduction.json), [export-hashes.csv](qa/export-hashes.csv), [artifacts.json](qa/artifacts.json) and [artifact-checks.txt](qa/artifact-checks.txt). They record executed results; a missing file or nonzero exit means the corresponding check has not completed.

| Check | Required evidence |
|---|---|
| Empty-database restore | Bundled snapshot restores with `--exit-on-error` into a new database; dump hash recorded. |
| Source/copy equality | Both audit outputs identical: 13 table counts and 40 hard-failing controls plus financial totals/metadata. |
| KPI and diagnostic regression | Five date cases (2025, 2024, mid-2025, partial January 2024, full history); 38 KPI checks and 15 diagnostic checks per database/case; eight complete diagnostic fingerprints match per case. Full logs are `docs/qa/kpi-*` and `analysis-*`. |
| Excel input reproduction | Every freshly exported CSV matches the accepted manifest SHA-256, including default references and commercial/support filter controls. Only scratch output paths are rewritten; reference query logic is unchanged. |
| Native numeric reconciliation | 282 required main values, 131 captured extended values, 10 focused cohort samples and 138 refresh values (46 each ORIGINAL/ALTERNATE/RESTORED) match the current source controls. |
| Workbook preservation | SHA-256 remains `72dce81f82b2d1c0aefa77201652ceaf81b33e9d8521552ce45a2bd22a35e047`, identical to the earlier CRC-validated, macro-free workbook with embedded model. Byte identity preserves that file-integrity evidence. |
| Optimization evidence | 54 saved measured plans and 18 summary rows retained; original report contains exact-equivalence, plan/buffer, range and index tradeoff evidence. Benchmarks are not rerun as a prerequisite for final reconciliation. |
| Documentation and hygiene | Source dictionary covers 13 tables; ERD covers 31 FKs; local documentation links resolve; no excluded local paths or credential-like files tracked/proposed; heuristic credential scan passes; `git diff --check` passes. |

The rechecked native results are saved Phase 4 captures, not newly executed Excel actions. There are **276 omitted exhaustive cohort cases**, which remain untested; do not relabel the interrupted run complete. Numeric coverage does not establish a new visual inspection, new refresh recovery or new save/reopen result. Final layout repairs, reopen and missing-folder recovery retain their **operator-reported** attribution in the [Phase 4 handoff](excel/phase4-handoff.md). This deliberately preserves accepted formatting and avoids exhaustive Verify or dashboard rebuild. Screenshots in README are authentic earlier native captures, not claims about newly captured final layout.

## Documentation completeness and boundaries

[README](../README.md) presents the business problem, numerical findings, source ERD, restore/reproduction workflow, dashboard screenshots and synthetic-data limitations. [Data dictionary](data_dictionary.md) and [KPI dictionary](kpi_dictionary.md) separate storage fields from reporting definitions; the [SQL catalog](sql_catalog.md) maps every query to purpose, sources and assumptions. The [optimization report](optimization_report.md), [recommendations](recommendations.md), [Excel model](excel_model.md) and [refresh/recovery guide](refresh.md) complete the analytical and operational handoff.

The credential check scans supported text formats in current tracked and proposed deliverables for private keys, common token forms, credential-bearing database URIs and literal password/token assignments. It is a heuristic, not an exhaustive secret audit of binaries or Git history. Authentication examples use container-local usernames and disposable trust configuration without a password. Local user paths are environment-specific paths, not credentials; set pFolder for the destination machine. Generated scratch CSVs and local status/instruction directories stay excluded.

Phase completion requires both runners and the Git whitespace check to pass. No production data, causal intervention effect, legal debt balance, full exhaustive Excel coverage or new UI acceptance is inferred from these checks.
