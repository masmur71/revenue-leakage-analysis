# Refresh and recovery guide

The delivered [workbook](../excel/RevenueLeakage.xlsx) is accepted and manually formatted. Start from this file. The [Phase 4 record](excel/phase4-handoff.md) distinguishes executed numeric checks from operator-reported final appearance, save/reopen and missing-folder recovery.

## Prepare validated inputs

1. Open PowerShell in the repository root with the PostgreSQL container running. Keep Excel closed or idle while exporting.
2. Run `./sql/05_excel/02_export.ps1`. Stop on any error; do not refresh from a partially written folder. Default source: read-only `bintel`; calendar 2025 and snapshot 2025-12-31.
3. Run `node excel/verify_exports.mjs`. Require exit code zero and PASS for schema, keys, relationships and source reconciliation. This tool does not execute M or DAX. The generated `data/excel/manifest.json` records hashes and dates. Preserve the accepted manifest until final QA comparison finishes.
4. Move the entire validated CSV folder together. Do not mix snapshots or files from separate exports. Generated `data/excel/` is excluded from Git; a fresh clone must export it.

## Refresh the existing workbook

1. Back up `excel/RevenueLeakage.xlsx`. Open it in Excel desktop for Windows with Data Model support; enable only the expected local data connections.
2. Open **Data → Get Data → Launch Power Query Editor**, select `pFolder`, and set its text value to the absolute CSV directory, for example `C:\Projects\leakageanalysis\data\excel`. In Advanced Editor, preserve parameter metadata and use an M text literal.
3. Apply the parameter, close the editor, and choose **Data → Refresh All**. Wait for all query/model work to finish. Inspect Queries & Connections for errors; row counts alone do not prove success.
4. Select year 2025 and clear branch/outlet/SKU/channel filters. Compare to `SQLReference.csv`: net sales **15,762,528,589.51**, gross profit **3,851,071,463.51**, returns **210,279,535.21**, outstanding AR **3,396,761,237.52**. Compare money to two decimals and BLANK separately from zero.
5. Select Medan: net sales **862,259,449.78**. Clear branch filters and confirm national revenue returns. Target comparisons require whole months and no SKU/outlet filter. AR/RFM snapshots do not move with operating-period slicers.
6. Check all five sheets for errors, clipped labels and charts. Save, close, reopen, then repeat the national/Medan spot check. Record new evidence separately from historical Phase 4 acceptance.

Do not run Build/Finish or exhaustive Phase4_Verify on the formatted deliverable. The `.xlsx` has no embedded VBA; separate `.bas` helpers document initial construction and are not required for refresh. Relationship/filter rules are in the [model guide](excel_model.md).

## Failure recovery

If the folder is missing, schema/type checks fail, or refresh is interrupted, stop. Cached totals are not newly refreshed results. Restore the complete validated folder and original `pFolder` value (or `original-pFolder.m` from a prior optional runner), then refresh again. Check query errors, national/Medan totals and save/reopen before acceptance. If recovery fails, close without saving and reopen the backup. Never replace missing values with zero or change KPI formulas to force reconciliation.

Export reads independent sale/return events in one repeatable-read snapshot. Changing `AsOfDate` requires a coherent new export and refresh; supported history ends on 2025-12-31. For initial model construction only, use the [installation guide](panduan_pemasangan_excel.md) and its native acceptance procedure. Regenerating M through build_contract resets folder paths and is unnecessary for the delivered workbook.
