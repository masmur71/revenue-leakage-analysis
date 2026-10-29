# Verification

Final QA completed on **30 October 2026**. The bundled synthetic dump restored into a new `leakage_qa_phase6` database; the original `bintel` source remained read-only.

| Check | Result |
|---|---|
| Restore and source/copy audit | 13 table counts and 40 controls matched |
| KPI and analysis regression | Five date scenarios matched: 38 KPI checks, 15 diagnostic checks and eight diagnostic fingerprints per case |
| Excel source reproduction | All 20 exported CSV hashes matched the accepted manifest |
| Saved native Excel captures | 282 main, 131 extended, 10 cohort samples and 138 refresh values reconciled |
| Workbook integrity | Unchanged, macro-free workbook with embedded Data Model; [integrity record](excel/final-workbook-integrity.json) |
| Optimization | 54 measured plans retained with summary and plan excerpts |

Reproduce using the [setup guide](setup.md). Runners regenerate detailed audit, regression and capture-comparison logs locally; these intermediate files are intentionally excluded from Git. The source [KPI contract](kpi_dictionary.md) and original reference SQL remain authoritative.

The native checks recompare saved Phase 4 captures; they are not a new Excel refresh or UI test. Final appearance, save/reopen and missing-folder recovery were operator-reported. The interrupted exhaustive run omitted **276 cohort cases**, which remain untested. No exhaustive Verify or dashboard rebuild was repeated.

Workbook SHA-256: `72dce81f82b2d1c0aefa77201652ceaf81b33e9d8521552ce45a2bd22a35e047`.

Current deliverables were checked for excluded local paths, broken documentation links and common credential patterns. Credential scanning is heuristic and does not establish absence of secrets in every binary or historical commit.
