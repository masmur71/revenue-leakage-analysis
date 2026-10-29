# Refresh and recovery

Start with the delivered [workbook](../excel/RevenueLeakage.xlsx), backed up before changing connections. Keep Excel idle while generating inputs.

1. Run `./sql/05_excel/02_export.ps1`, then `node excel/verify_exports.mjs`. Stop on errors; never use a partially written CSV folder.
2. Open the workbook in Excel desktop for Windows. In **Power Query Editor**, set `pFolder` to the absolute `data/excel` directory containing the complete export.
3. Choose **Refresh All**, wait for completion, and check Queries & Connections for errors.
4. Select 2025 and clear other filters. Verify net sales **15,762,528,589.51**, gross profit **3,851,071,463.51**, and outstanding AR **3,396,761,237.52** against `SQLReference.csv`.
5. Filter Medan: net sales should be **862,259,449.78**. Clear the filter, confirm the national total, then save and reopen to repeat the spot check.

Keep the whole CSV export together when moving machines. Never mix snapshots. Default snapshot is 2025-12-31; AR/RFM snapshots do not change with operating-period slicers. Target comparisons require whole months and no SKU/outlet filter. The [KPI dictionary](kpi_dictionary.md) defines the filter and date rules; the [model contract](../excel/model_contract.json) and [relationships](../excel/relationships.csv) define the workbook inputs.

If refresh fails, restore the complete validated folder and original `pFolder`, then refresh and recheck totals. Cached totals are not proof of a successful refresh. If recovery fails, close without saving and reopen the backup. Never replace missing values with zero to force reconciliation.

Preserve manual formatting. Do not run Build/Finish or exhaustive Phase4_Verify for routine refresh. The delivered `.xlsx` contains no VBA. See [setup](setup.md) for restoration and [verification](final_qa.md) for acceptance coverage.
