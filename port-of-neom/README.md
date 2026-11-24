# Port of NEOM – Excel data pack and Power BI report

| File | What it is |
|---|---|
| `Port_of_NEOM_Data.xlsx` | The data, in 19 Excel Tables: an illustrative commercial, operations and finance model plus sourced public facts. |
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

## Two kinds of data

**Illustrative model (orange tabs, first five report pages).** The port publishes no commercial,
operating or financial results, so these are generated: synthetic monthly volumes shaped to the
public picture, fictional customers, assumed tariffs and cost drivers. Revenue, costs and the
forecast are Excel formulas driven by the `Assumptions` sheet. **None of it is an actual Port of
NEOM figure**, and every sheet and page says so.

**Public facts (uncoloured tabs, last four report pages).** Sourced figures; see the `Sources` sheet.

## Report pages

| Page | Content |
|---|---|
| Executive Summary | Revenue, EBITDA, margin, tonnes, calls, berth occupancy; revenue vs budget; tonnes by cargo type; top customers |
| Commercial | Revenue by stream, customer and cargo type; discounts; revenue per tonne; customer league table |
| Operations | Vessel calls, time at berth, berth occupancy, crane productivity, truck turn time, anchorage wait, injury rate |
| Finance | Income statement lines actual vs budget, operating costs by account, EBITDA trend, capex by project |
| Forecast 2027-2031 | Container volume, revenue and EBITDA under Base, High and Low scenarios |
| Public Facts (4 pages) | Port profile, infrastructure and investment, vessel traffic and land bridge, timeline and sources |

Monthly data runs January 2023 to October 2026. Each model page has a Year slicer.

## Using real data

Overwrite the inputs and the formulas follow: `Operations` (Volume, VesselCalls, BerthHours),
the first six `OpsKPI` columns, `Customers`, `CargoTypes` rates, `Capex`, and the `Finance`
budget rows. To load ledger actuals, paste values over the `Finance` "Actual" formulas. Keep the
column headers and table names unchanged, then refresh Power BI.

## What was and was not checked

- The workbook was recalculated: 5,204 formulas, no errors. Finance revenue ties to the
  Commercial sheet to zero, and the Excel results match an independent Python calculation.
- Every report JSON file validates against Microsoft's published PBIR schemas.
- The semantic model loads in Microsoft's TMDL library (19 tables, 17 relationships, 71 measures).
- Every field a visual uses exists in the model; every model column exists in the workbook with a
  loadable type; every relationship key is unique and complete.
- **The project has not been opened in Power BI Desktop**, which only runs on Windows. Layout and
  formatting are unseen; expect to nudge sizes and fonts on first open.

## About the public facts

All sourced figures come from public reporting found by web search on 24 November 2026.

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
