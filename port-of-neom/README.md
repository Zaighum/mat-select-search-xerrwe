# Port of NEOM – Excel data pack and Power BI report

| File | What it is |
|---|---|
| `Port_of_NEOM_Data.xlsx` | The data, in nine Excel Tables, plus a Summary sheet with key figures and two charts. |
| `powerbi/PortOfNEOM.pbip` | Power BI project (report + semantic model) that loads the workbook. |
| `PortOfNEOM_PowerBI.zip` | The workbook and the `powerbi` folder in one download. |
| `build/` | The Python scripts that generate both. |

## Open the report

You need Power BI Desktop on Windows (a 2025 or later release).

1. Unzip `PortOfNEOM_PowerBI.zip`. Keep the `powerbi` folder intact.
2. Copy `Port_of_NEOM_Data.xlsx` to `C:\PortOfNEOM\Port_of_NEOM_Data.xlsx`.
   To keep it elsewhere, change the path in step 4.
3. Double-click `powerbi\PortOfNEOM.pbip`.
   If Desktop does not offer to open `.pbip` files, switch on
   **File > Options > Preview features > Power BI Project (.pbip) save option** and
   **Store reports using enhanced metadata format (PBIR)**, then restart.
4. If the workbook is not at the default path: **Home > Transform data > Edit parameters**,
   set `ExcelFilePath`.
5. Click **Refresh**. The project ships without data, so the visuals are empty until you do.
6. **File > Save as** and pick `.pbix` if you want a single file to share.

## Report pages

1. **Overview** – Terminal 1 capacity, quay, depth, cranes, investment, status, capacity plan chart, key-facts table.
2. **Infrastructure and Investment** – equipment by type, stated investment, reported contracts.
3. **Traffic and Trade Corridor** – vessel calls per month before and after the war began, the Europe–Egypt–NEOM–GCC/Iraq land bridge.
4. **Timeline and Sources** – dated events with category and status slicers, and the source list with links.

## What was and was not checked

- The workbook was recalculated: 45 formulas, no errors.
- Every report JSON file validates against Microsoft's published PBIR schemas.
- The semantic model loads in Microsoft's TMDL library (9 tables, 8 relationships, 22 measures).
- Every field a visual uses exists in the model, and every model column exists in the workbook.
- **The project has not been opened in Power BI Desktop**, which only runs on Windows. Layout and
  formatting are therefore unseen; expect to nudge sizes and fonts on first open.

## About the data

All figures come from public reporting found by web search on 24 November 2026; the `Sources`
sheet lists each one. Read the `README` sheet in the workbook before quoting anything. In short:

- No official throughput statistics for the port were found. Capacity figures are plans or operator claims.
- Terminal 1 was announced to open on 10 November 2026; no confirmation that it opened was found.
- Vessel-call counts come from one analytics vendor (Windward).
- The three investment figures come from different statements and are not a time series.
- Figures were taken from search-result extracts, not from the pages themselves. Check a figure
  against its URL before using it externally.

## Rebuild

```
python build/build_excel.py Port_of_NEOM_Data.xlsx
python <xlsx-skill>/scripts/recalc.py Port_of_NEOM_Data.xlsx      # caches formula values
python build/build_pbip.py powerbi build/CY24SU10.json
```
