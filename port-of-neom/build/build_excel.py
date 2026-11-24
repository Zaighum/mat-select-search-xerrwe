"""Builds Port_of_NEOM_Data.xlsx - the data source for the Power BI report.

All figures come from public reporting found by web search (see the Sources sheet).
Run:  python build_excel.py <output.xlsx>
"""
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

FONT = "Arial"
NAVY = "0B2545"
TEAL = "13A89E"
HDR_ROW = 4  # header row of every data table

# ----------------------------------------------------------------------------
# DATA  (every row carries a SourceID that resolves on the Sources sheet)
# ----------------------------------------------------------------------------
SOURCES = [
    # id, publisher, title, published, url, note
    ("S01", "Port of NEOM (official)", "Port of NEOM - home page", "Accessed via search, Nov 2026",
     "https://www.portofneom.com/en-us", "Operator's own claims"),
    ("S02", "Port of NEOM (official)", "Terminal 1", "Accessed via search, Nov 2026",
     "https://www.portofneom.com/en-us/terminal-1", "Operator's own claims; wording on launch year changed in Aug 2026 per AGBI"),
    ("S03", "AGBI", "Port of Neom to open first terminal next month", "Oct 2026",
     "https://www.agbi.com/logistics/2026/10/port-of-neom-to-open-first-terminal-next-month/",
     "Also published as .../port-of-neom-to-open-container-terminal-next-month/"),
    ("S04", "Inspenet", "The port of Neom will open its container terminal on November 10th", "Oct 2026",
     "https://inspenet.com/en/news/the-port-of-neom-container-terminal/", ""),
    ("S05", "Breakbulk", "Port of NEOM: A Benchmark for the Maritime Industry", "c. 2024 (undated in search result)",
     "https://breakbulk.com/articles/port-of-neom-a-benchmark-for-the-industry", "Interview-based; plans, not outcomes"),
    ("S06", "Project Cargo Weekly", "Port of NEOM (interview)", "29 May 2024",
     "https://www.projectcargo-weekly.com/2024/05/29/port-of-neom/", "Interview-based; plans, not outcomes"),
    ("S07", "Trade Arabia", "Saudi Port of NEOM opens following $2bn investment", "2023",
     "https://www.tradearabia.com/News/274807/Saudi-Port-of-NEOM-opens-following-%242bn-investment", ""),
    ("S08", "Saudi Press Agency", "Port of NEOM Open for Business", "2023",
     "https://www.spa.gov.sa/en/c74b80d4b0q", "Official announcement"),
    ("S09", "Windward", "War-Diverted Traffic Is Reshaping NEOM's Port. What Happens If Hormuz Reopens?", "c. Oct 2026",
     "https://windward.ai/blog/war-diverted-traffic-is-reshaping-neoms-port/",
     "Single vendor's vessel-tracking dataset; not official port statistics"),
    ("S10", "Arabian Business (Reuters-sourced)", "Saudi Arabia's NEOM port gains as Iran war forces Gulf firms to find new trade routes", "2026",
     "https://www.arabianbusiness.com/business/transport/saudi-arabias-neom-port-gains-as-iran-war-forces-gulf-firms-to-find-new-trade-routes", ""),
    ("S11", "NEOM newsroom", "Port of NEOM and Partners Enable European Multimodal Land Bridge", "Apr 2026",
     "https://www.neom.com/en-us/newsroom/Port-of-NEOM-and-Partners-Enable-European-Multimodal-Land-Bridge-Advancing-Regional-Cargo-Connectivity-",
     "Official announcement"),
    ("S12", "NEOM newsroom", "Port of NEOM and CEVA Almajdouie Logistics collaborate to enhance connectivity", "Oct 2026",
     "https://www.neom.com/en-us/newsroom/Port-of-NEOM-CEVA-Almajdouie-Logistics-collaborate-to-enhance-connectivity",
     "MoU - exploratory, not a measured result"),
    ("S13", "NEOM (X / NEOM Now)", "NEOM Now Edition 5: Port of NEOM, Connecting Saudi to the World", "2026",
     "https://x.com/NEOM/status/2047289147392475413", "Promotional"),
    ("S14", "The National", "How Saudi Arabia is repositioning Neom as key logistics hub", "17 Apr 2026",
     "https://www.thenationalnews.com/business/economy/2026/04/17/how-saudi-arabia-is-repositioning-neom-as-key-logistics-hub-and-economic-anchor/", ""),
    ("S15", "InfraPPP World", "US$800 million Contract Awarded for building Port of Neom", "Undated in search result",
     "https://www.infrapppworld.com/update/us800-million-contract-awarded-for-building-port-of-neom-in-saudi-arabia",
     "Headline figure only; contractor and scope not captured"),
    ("S16", "Construction Business News ME", "Port of Neom, Saudi Arabia", "Undated in search result",
     "https://www.cbnme.com/top-ports-in-the-middle-east/port-of-neom-saudi-arabia/", "Reports early plan of >2m TEU, start-up 2025"),
    ("S17", "Hellenic Shipping News", "War-Diverted Traffic Is Reshaping NEOM's Port (syndicated copy)", "c. Oct 2026",
     "https://www.hellenicshippingnews.com/war-diverted-traffic-is-reshaping-neoms-port-what-happens-if-hormuz-reopens/",
     "Syndicated Windward text; gives 25 bulk calls where Windward's page gives 24"),
]

# id, category, attribute, text value, numeric value, unit, source, note
PROFILE = [
    ("P01", "Identity", "Location", "Oxagon, NEOM - Red Sea coast, north-west Saudi Arabia", None, "", "S01", ""),
    ("P02", "Identity", "Former name", "Duba Port", None, "", "S03", "Operation handed from Mawani to NEOM in 2022; reopened as Port of NEOM in 2023"),
    ("P03", "Identity", "Terminal 1 status", "Opening announced for 10 Nov 2026 - not independently confirmed", None, "", "S03", "No post-opening report found as of 24 Nov 2026"),
    ("P04", "Terminal 1", "Quay wall length", "900 m", 900, "m", "S02", ""),
    ("P05", "Terminal 1", "Vessel berths", "2", 2, "berths", "S02", ""),
    ("P06", "Marine access", "Draft / basin depth", "18.5 m", 18.5, "m", "S01", ""),
    ("P07", "Marine access", "Access channel width", "550 m", 550, "m", "S01", ""),
    ("P08", "Terminal 1", "Phase 1 container capacity", "700,000 TEU", 700000, "TEU / year", "S03", "AGBI reading of government data"),
    ("P09", "Terminal 1", "Full build-up container capacity", "1.5 million TEU", 1500000, "TEU / year", "S02", "Target year 2030 per AGBI (S03)"),
    ("P10", "Terminal 1", "First full year of operations", "2027", 2027, "year", "S02", ""),
    ("P11", "General cargo", "General cargo and bulk capacity target", "20 million freight tonnes", 20000000, "freight tonnes / year", "S05", "Stated target, also in S06"),
    ("P12", "General cargo", "Project and bulk cargo scale cited by NEOM", "9 million tonnes", 9000000, "tonnes", "S13", "Promotional figure; basis not stated"),
    ("P13", "Market context", "Saudi ports container volume (prior year, implied)", "about 9 million TEU", 9000000, "TEU / year", "S03", "AGBI: 700,000 TEU is roughly 8% of Saudi port TEU last year"),
    ("P14", "Market context", "Trucks arriving at Safaga for the Red Sea crossing", "about 1,000 per day", 1000, "trucks / day", "S10", "Reuters-sourced report, 2026"),
]

# id, date, precision, event, category, status, source
TIMELINE = [
    ("T01", date(2022, 1, 1), "Year", "Operation of Duba Port handed over from Mawani (Saudi Ports Authority) to NEOM", "Governance", "Completed", "S03"),
    ("T02", date(2023, 1, 1), "Year", "Port reopens as 'Port of NEOM' - declared open for business", "Operations", "Completed", "S08"),
    ("T03", date(2023, 1, 1), "Year", "Investment of over SAR 7.5 billion (about US$2bn) announced; Terminal 1 targeted for 2025", "Investment", "Completed", "S07"),
    ("T04", date(2024, 5, 29), "Day", "Managing director interview: investment over SAR 14 billion; 1.5m TEU terminal planned", "Investment", "Reported", "S06"),
    ("T05", date(2025, 1, 1), "Year", "Original Terminal 1 opening target (not met)", "Terminal 1", "Missed", "S05"),
    ("T06", date(2025, 1, 1), "Year", "Intermodal pilot: Safaga (Egypt) - Port of NEOM - Arar border - Iraq", "Trade corridor", "Completed", "S11"),
    ("T07", date(2026, 4, 1), "Month", "European multimodal land bridge enabled: Europe - Egypt - Port of NEOM - GCC / Iraq", "Trade corridor", "Completed", "S11"),
    ("T08", date(2026, 8, 1), "Month", "Port website wording changes to Terminal 1 launching 'from 2027 onwards'", "Terminal 1", "Reported", "S03"),
    ("T09", date(2026, 9, 1), "Month", "RoRo calls reach 30 in the month (pre-war average 17), mostly Safaga shuttle", "Operations", "Reported", "S09"),
    ("T10", date(2026, 10, 1), "Month", "MoU with CEVA Almajdouie Logistics on integrated supply-chain solutions", "Trade corridor", "Completed", "S12"),
    ("T11", date(2026, 11, 10), "Day", "Announced Terminal 1 opening date (Phase 1: 700,000 TEU)", "Terminal 1", "Unconfirmed", "S03"),
    ("T12", date(2027, 1, 1), "Year", "First full year of Terminal 1 operations", "Terminal 1", "Planned", "S02"),
    ("T13", date(2030, 1, 1), "Year", "Terminal 1 to reach 1.5 million TEU annual capacity", "Terminal 1", "Planned", "S03"),
]

# id, statement year, milestone, target year, capacity TEU, plan status, source, note
CAPACITY = [
    ("C01", 2023, "Early plan", 2025, 2000000, "Superseded", "S16", "Reported as 'more than 2 million TEU'"),
    ("C02", 2024, "T1 planned capacity", 2025, 1500000, "Superseded", "S05", "Capacity kept, opening date slipped"),
    ("C03", 2026, "T1 Phase 1", 2026, 700000, "Current", "S03", "At opening"),
    ("C04", 2026, "T1 full build-up", 2030, 1500000, "Current", "S03", "Port site (S02) gives 1.5m TEU without the 2030 date"),
]

# id, equipment, code, units, supplier, power/automation, source
EQUIPMENT = [
    ("E01", "Ship-to-shore gantry cranes", "STS", 10, "ZPMC", "Automated / remote-controlled", "S05"),
    ("E02", "Electric rubber-tyred gantry cranes", "eRTG", 30, "ZPMC", "Electric", "S05"),
    ("E03", "Automated rail-mounted gantry cranes", "ARMG", 6, "ZPMC", "Automated", "S05"),
    ("E04", "Mobile harbour cranes", "MHC", 10, "Liebherr (Saudi Liebherr)", "Mobile", "S05"),
]

# id, counterparty, scope, currency, stated value (millions, original currency), source, note
CONTRACTS = [
    ("K01", "ZPMC (with Siemens for automation)", "10 STS + 30 eRTG + 6 ARMG cranes", "SAR", 1000, "S05", "Stated as 'over SAR 1 billion'"),
    ("K02", "Saudi Liebherr", "10 mobile harbour cranes", "SAR", 200, "S05", "Stated as 'over SAR 200 million'"),
    ("K03", "Not captured", "Port construction contract", "USD", 800, "S15", "Headline figure only"),
]

# id, statement, year, SAR bn, source, note
INVESTMENT = [
    ("I01", "2023 opening announcement", 2023, 7.5, "S08", "'Over SAR 7.5 billion'; reported as about US$2bn (S07)"),
    ("I02", "Breakbulk interview (c. 2024)", 2024, 12, "S05", "'Investment to date is over SAR 12 billion'"),
    ("I03", "Project Cargo Weekly, May 2024", 2024, 14, "S06", "'Over SAR 14 billion (close to 4 billion USD)'"),
]

# id, vessel type, period group, period detail, calls, months, stated avg/month, source, note
VESSELS = [
    ("V01", "RoRo", "Pre-war", "Pre-war monthly average", None, None, 17, "S09", "Average stated directly; underlying counts not given"),
    ("V02", "RoRo", "After war began", "September 2026", 30, 1, None, "S09", "Almost all on the Safaga shuttle"),
    ("V03", "Bulk", "Pre-war", "7 months before the war", 1, 7, None, "S09", ""),
    ("V04", "Bulk", "After war began", "5 months after the war began", 24, 5, None, "S09", "Syndicated copy (S17) says 25"),
    ("V05", "Container", "Pre-war", "5 months before the war", 7, 5, None, "S09", "Mostly regional feeders"),
    ("V06", "Container", "After war began", "March - July 2026", 2, 5, None, "S09", "Zero calls since mid-July 2026"),
]

# id, sequence, node, country, role, onward mode, lat, lon, source
CORRIDOR = [
    ("R01", 1, "Europe (origin markets)", "Europe", "Origin", "Truck / RoRo to Egypt", None, None, "S11"),
    ("R02", 2, "Port of Safaga", "Egypt", "Red Sea departure port", "RoRo ferry shuttle to Port of NEOM", 26.73, 33.94, "S11"),
    ("R03", 3, "Port of NEOM", "Saudi Arabia", "Gateway / transfer hub", "Truck across northern Saudi Arabia", 27.56, 35.53, "S11"),
    ("R04", 4, "Arar border crossing", "Saudi Arabia / Iraq", "Land border", "Truck into Iraq", 30.98, 41.02, "S11"),
    ("R05", 5, "Iraq", "Iraq", "Destination", "", None, None, "S11"),
    ("R06", 5, "UAE", "United Arab Emirates", "Destination", "", None, None, "S11"),
    ("R07", 5, "Kuwait", "Kuwait", "Destination", "", None, None, "S11"),
    ("R08", 5, "Oman", "Oman", "Destination", "", None, None, "S11"),
]

# ----------------------------------------------------------------------------
# HELPERS
# ----------------------------------------------------------------------------
thin = Side(style="thin", color="D9D9D9")


def base(ws):
    ws.sheet_view.showGridLines = False


def title(ws, text, sub):
    ws["A1"] = text
    ws["A1"].font = Font(name=FONT, size=14, bold=True, color=NAVY)
    ws["A2"] = sub
    ws["A2"].font = Font(name=FONT, size=9, italic=True, color="595959")


def data_sheet(wb, name, table, headers, rows, widths, sub, formats=None):
    """Writes a titled sheet with one Excel Table starting at row HDR_ROW."""
    ws = wb.create_sheet(name)
    base(ws)
    title(ws, name, sub)
    for c, h in enumerate(headers, 1):
        ws.cell(HDR_ROW, c, h).font = Font(name=FONT, bold=True, color="FFFFFF")
    for r, row in enumerate(rows, HDR_ROW + 1):
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v)
            cell.font = Font(name=FONT, size=10)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if formats and headers[c - 1] in formats:
                cell.number_format = formats[headers[c - 1]]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ref = f"A{HDR_ROW}:{get_column_letter(len(headers))}{HDR_ROW + len(rows)}"
    t = Table(displayName=table, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(t)
    ws.freeze_panes = ws.cell(HDR_ROW + 1, 1)
    return ws


def col(headers, name):
    return get_column_letter(headers.index(name) + 1)


# ----------------------------------------------------------------------------
# BUILD
# ----------------------------------------------------------------------------
def build(path):
    wb = Workbook()

    # ---- README ------------------------------------------------------------
    ws = wb.active
    ws.title = "README"
    base(ws)
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 110
    ws["A1"] = "Port of NEOM - data pack for Power BI"
    ws["A1"].font = Font(name=FONT, size=16, bold=True, color=NAVY)
    lines = [
        ("Purpose", "Source data for the Power BI project 'PortOfNEOM.pbip'. Each data sheet holds one Excel Table (tbl...) that Power BI loads."),
        ("Compiled", "24 November 2026, from public reporting found by web search. Nothing here is from NEOM's internal systems."),
        ("How to read it", "Every row has a SourceID. Look it up on the Sources sheet for the publisher, title and URL."),
        ("Formula columns", "Black-on-white cells with formulas (USD conversions, averages, shares) recalculate from the inputs. The SAR/USD rate is on the Assumptions sheet."),
        ("Caveat 1 - no official throughput", "No published TEU, tonnage or vessel-call statistics for the Port of NEOM were found. Capacity figures are plans or operator claims, not measured volumes."),
        ("Caveat 2 - Terminal 1 status", "Terminal 1 was announced to open on 10 Nov 2026. No report confirming that it opened was found as of 24 Nov 2026, so it is marked 'Unconfirmed'."),
        ("Caveat 3 - vessel calls", "Vessel-call counts come from one analytics vendor (Windward) and cover unequal periods; they are normalised to calls per month for comparison."),
        ("Caveat 4 - investment", "The three investment figures come from different statements and dates and are not a consistent time series."),
        ("Caveat 5 - coordinates", "Latitude/longitude on the Corridor sheet are approximate, added for mapping only, and are not from the cited sources."),
        ("Caveat 6 - extraction", "Figures were taken from search-result extracts; the source pages could not be opened directly from the build environment. Check a figure against its URL before quoting it externally."),
    ]
    for i, (k, v) in enumerate(lines, 3):
        ws.cell(i, 1, k).font = Font(name=FONT, bold=True, size=10)
        c = ws.cell(i, 2, v)
        c.font = Font(name=FONT, size=10)
        c.alignment = Alignment(wrap_text=True, vertical="top")

    # ---- Assumptions -------------------------------------------------------
    wa = wb.create_sheet("Assumptions")
    base(wa)
    title(wa, "Assumptions", "Blue = hardcoded input. Change it here and every conversion updates.")
    wa["A4"], wa["B4"], wa["C4"] = "SAR per USD", 3.75, "Official Saudi riyal peg to the US dollar (fixed since 1986)."
    wa["A4"].font = Font(name=FONT, bold=True)
    wa["B4"].font = Font(name=FONT, color="0000FF", bold=True)
    wa["B4"].fill = PatternFill("solid", fgColor="FFFF00")
    wa["B4"].number_format = "0.00"
    wa["C4"].font = Font(name=FONT, size=10, italic=True)
    wa.column_dimensions["A"].width = 22
    wa.column_dimensions["B"].width = 12
    wa.column_dimensions["C"].width = 70
    FX = "Assumptions!$B$4"

    # ---- data sheets -------------------------------------------------------
    h_profile = ["ProfileID", "Category", "Attribute", "Value", "NumericValue", "Unit", "SourceID", "Note"]
    data_sheet(wb, "PortProfile", "tblPortProfile", h_profile, PROFILE, [10, 16, 44, 54, 14, 20, 10, 60],
               "Key facts about the port. NumericValue is blank for text-only facts.", {"NumericValue": "#,##0.0"})

    h_time = ["EventID", "EventDate", "DatePrecision", "Event", "Category", "Status", "SourceID"]
    data_sheet(wb, "Timeline", "tblTimeline", h_time, TIMELINE, [9, 13, 14, 84, 16, 14, 10],
               "EventDate is the first day of the stated period; DatePrecision says how exact the source was.",
               {"EventDate": "dd-mmm-yyyy"})

    h_cap = ["CapacityID", "StatementYear", "Milestone", "TargetYear", "CapacityTEU", "PlanStatus", "SourceID", "Note", "ShareOfSaudiTEU"]
    rows = [r + (None,) for r in CAPACITY]
    wc = data_sheet(wb, "CapacityPlan", "tblCapacityPlan", h_cap, rows, [11, 14, 24, 12, 14, 13, 10, 52, 17],
                    "Container capacity plans as announced over time. ShareOfSaudiTEU = capacity / implied Saudi port volume (PortProfile P13).",
                    {"CapacityTEU": "#,##0", "ShareOfSaudiTEU": "0.0%"})
    saudi = f"INDEX(PortProfile!$E:$E,MATCH(\"P13\",PortProfile!$A:$A,0))"
    for r in range(HDR_ROW + 1, HDR_ROW + 1 + len(CAPACITY)):
        wc[f"I{r}"] = f"=E{r}/{saudi}"
        wc[f"I{r}"].number_format = "0.0%"
        wc[f"I{r}"].font = Font(name=FONT, size=10)

    h_eq = ["EquipmentID", "Equipment", "TypeCode", "Units", "Supplier", "Technology", "SourceID"]
    data_sheet(wb, "Equipment", "tblEquipment", h_eq, EQUIPMENT, [12, 38, 10, 8, 26, 30, 10],
               "Cargo-handling equipment reported as ordered for the port.", {"Units": "0"})

    h_con = ["ContractID", "Counterparty", "Scope", "Currency", "StatedValueM", "SourceID", "Note", "ValueSARm", "ValueUSDm"]
    rows = [r + (None, None) for r in CONTRACTS]
    wk = data_sheet(wb, "Contracts", "tblContracts", h_con, rows, [11, 34, 36, 10, 13, 10, 34, 12, 12],
                    "Reported contract values, in millions. Values stated as 'over X' are held at X. Conversions use the Assumptions rate.",
                    {"StatedValueM": "#,##0"})
    for r in range(HDR_ROW + 1, HDR_ROW + 1 + len(CONTRACTS)):
        wk[f"H{r}"] = f'=IF(D{r}="SAR",E{r},E{r}*{FX})'
        wk[f"I{r}"] = f'=IF(D{r}="USD",E{r},E{r}/{FX})'
        for c in "HI":
            wk[f"{c}{r}"].number_format = "#,##0"
            wk[f"{c}{r}"].font = Font(name=FONT, size=10)

    h_inv = ["InvestmentID", "Statement", "StatementYear", "AmountSARbn", "SourceID", "Note", "AmountUSDbn"]
    rows = [r + (None,) for r in INVESTMENT]
    wi = data_sheet(wb, "Investment", "tblInvestment", h_inv, rows, [13, 34, 14, 13, 10, 60, 13],
                    "Cumulative investment as stated at different times ('over X' held at X). Not a consistent time series.",
                    {"AmountSARbn": "0.0"})
    for r in range(HDR_ROW + 1, HDR_ROW + 1 + len(INVESTMENT)):
        wi[f"G{r}"] = f"=D{r}/{FX}"
        wi[f"G{r}"].number_format = "0.00"
        wi[f"G{r}"].font = Font(name=FONT, size=10)

    h_ves = ["CallID", "VesselType", "PeriodGroup", "PeriodDetail", "Calls", "Months", "StatedAvgPerMonth", "SourceID", "Note", "AvgCallsPerMonth"]
    rows = [r + (None,) for r in VESSELS]
    wv = data_sheet(wb, "VesselCalls", "tblVesselCalls", h_ves, rows, [8, 12, 16, 30, 8, 9, 18, 10, 48, 17],
                    "Vessel calls per Windward (S09). AvgCallsPerMonth = Calls / Months, or the stated average where counts were not given.",
                    {"Calls": "0", "Months": "0"})
    for r in range(HDR_ROW + 1, HDR_ROW + 1 + len(VESSELS)):
        wv[f"J{r}"] = f"=IF(AND(ISNUMBER(E{r}),ISNUMBER(F{r}),F{r}>0),E{r}/F{r},G{r})"
        wv[f"J{r}"].number_format = "0.0"
        wv[f"J{r}"].font = Font(name=FONT, size=10)

    h_cor = ["NodeID", "Sequence", "Node", "Country", "Role", "OnwardMode", "Latitude", "Longitude", "SourceID"]
    data_sheet(wb, "Corridor", "tblCorridor", h_cor, CORRIDOR, [8, 10, 26, 22, 24, 38, 10, 10, 10],
               "Europe - Egypt - Port of NEOM - GCC / Iraq land bridge (S11). Coordinates are approximate and not from the source.",
               {"Latitude": "0.00", "Longitude": "0.00"})

    h_src = ["SourceID", "Publisher", "Title", "Published", "URL", "Note"]
    wsrc = data_sheet(wb, "Sources", "tblSources", h_src, SOURCES, [9, 30, 62, 28, 70, 60],
                      "Every source used. 'Published' is as shown in search results; undated items are flagged.")
    for r in range(HDR_ROW + 1, HDR_ROW + 1 + len(SOURCES)):
        c = wsrc[f"E{r}"]
        c.hyperlink = c.value
        c.font = Font(name=FONT, size=10, color="0563C1", underline="single")

    # ---- Summary -----------------------------------------------------------
    s = wb.create_sheet("Summary", 1)
    base(s)
    title(s, "Port of NEOM - summary", "All figures are formulas over the data sheets. Capacity figures are plans, not measured throughput.")
    s.column_dimensions["A"].width = 46
    s.column_dimensions["B"].width = 18
    s.column_dimensions["C"].width = 20
    s.column_dimensions["D"].width = 60

    def prof(pid):
        return f'=INDEX(PortProfile!$E:$E,MATCH("{pid}",PortProfile!$A:$A,0))'

    def cap(cid, colname):
        return f'=INDEX(CapacityPlan!${col(h_cap, colname)}:${col(h_cap, colname)},MATCH("{cid}",CapacityPlan!$A:$A,0))'

    def ves(vid):
        return f'=INDEX(VesselCalls!$J:$J,MATCH("{vid}",VesselCalls!$A:$A,0))'

    kpis = [
        ("Terminal 1", None, None, None),
        ("Phase 1 container capacity", cap("C03", "CapacityTEU"), "TEU / year", "#,##0"),
        ("Full build-up capacity (target 2030)", cap("C04", "CapacityTEU"), "TEU / year", "#,##0"),
        ("Phase 1 as share of full build-up", "=B6/B7", "", "0.0%"),
        ("Phase 1 as share of Saudi port TEU (implied)", cap("C03", "ShareOfSaudiTEU"), "", "0.0%"),
        ("Quay wall", prof("P04"), "m", "#,##0"),
        ("Berths", prof("P05"), "", "0"),
        ("Draft / basin depth", prof("P06"), "m", "0.0"),
        ("Equipment and investment", None, None, None),
        ("Cranes reported ordered (all types)", "=SUM(Equipment!$D:$D)", "units", "0"),
        ("  of which ship-to-shore", '=SUMIFS(Equipment!$D:$D,Equipment!$C:$C,"STS")', "units", "0"),
        ("Crane contracts (stated minimum)", '=SUMIFS(Contracts!$H:$H,Contracts!$C:$C,"*cranes*")', "SAR million", "#,##0"),
        ("Highest stated cumulative investment", "=MAX(Investment!$D:$D)", "SAR billion", "0.0"),
        ("  in US dollars", "=B17/Assumptions!$B$4", "USD billion", "0.00"),
        ("Vessel traffic (Windward)", None, None, None),
        ("RoRo calls per month - pre-war", ves("V01"), "calls / month", "0.0"),
        ("RoRo calls per month - September 2026", ves("V02"), "calls / month", "0.0"),
        ("RoRo change vs pre-war", "=B21/B20-1", "", "0%"),
        ("Bulk calls per month - pre-war", ves("V03"), "calls / month", "0.0"),
        ("Bulk calls per month - after war began", ves("V04"), "calls / month", "0.0"),
        ("Container calls per month - pre-war", ves("V05"), "calls / month", "0.0"),
        ("Container calls per month - Mar-Jul 2026", ves("V06"), "calls / month", "0.0"),
    ]
    s["A4"], s["B4"], s["C4"] = "Measure", "Value", "Unit"
    for c in "ABC":
        s[f"{c}4"].font = Font(name=FONT, bold=True, color="FFFFFF")
        s[f"{c}4"].fill = PatternFill("solid", fgColor=NAVY)
    for i, (label, formula, unit, fmt) in enumerate(kpis, 5):
        a = s.cell(i, 1, label)
        if formula is None:
            a.font = Font(name=FONT, bold=True, color=NAVY)
            for c in range(1, 4):
                s.cell(i, c).fill = PatternFill("solid", fgColor="E8F4F3")
            continue
        a.font = Font(name=FONT, size=10)
        b = s.cell(i, 2, formula)
        b.font = Font(name=FONT, size=10, bold=True)
        b.number_format = fmt
        s.cell(i, 3, unit).font = Font(name=FONT, size=10, color="595959")
        for c in range(1, 4):
            s.cell(i, c).border = Border(bottom=thin)
    s["A29"] = "Terminal 1 status"
    s["A29"].font = Font(name=FONT, bold=True, color=NAVY)
    s["B29"] = '=INDEX(PortProfile!$D:$D,MATCH("P03",PortProfile!$A:$A,0))'
    s["B29"].font = Font(name=FONT, size=10, bold=True, color="C00000")

    # chart 1: capacity plan
    ch = BarChart()
    ch.type = "col"
    ch.title = "Container capacity plan by milestone (TEU / year)"
    ch.add_data(Reference(wc, min_col=5, min_row=HDR_ROW, max_row=HDR_ROW + len(CAPACITY)), titles_from_data=True)
    ch.set_categories(Reference(wc, min_col=3, min_row=HDR_ROW + 1, max_row=HDR_ROW + len(CAPACITY)))
    ch.legend = None
    ch.roundedCorners = False
    ch.title.overlay = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.numFmt = "#,##0"
    ch.height, ch.width = 8.5, 17
    s.add_chart(ch, "E4")

    # chart 2: vessel calls per month, by type and period (helper block, formulas)
    s["E24"], s["F24"], s["G24"] = "Vessel type", "Pre-war", "After war began"
    for c in "EFG":
        s[f"{c}24"].font = Font(name=FONT, bold=True, size=9)
    for i, vt in enumerate(["RoRo", "Bulk", "Container"], 25):
        s[f"E{i}"] = vt
        s[f"F{i}"] = f'=SUMIFS(VesselCalls!$J:$J,VesselCalls!$B:$B,$E{i},VesselCalls!$C:$C,F$24)'
        s[f"G{i}"] = f'=SUMIFS(VesselCalls!$J:$J,VesselCalls!$B:$B,$E{i},VesselCalls!$C:$C,G$24)'
        for c in "EFG":
            s[f"{c}{i}"].font = Font(name=FONT, size=9)
            s[f"{c}{i}"].number_format = "0.0"
    ch2 = BarChart()
    ch2.type = "col"
    ch2.title = "Average vessel calls per month, before and after the war began"
    ch2.add_data(Reference(s, min_col=6, max_col=7, min_row=24, max_row=27), titles_from_data=True)
    ch2.set_categories(Reference(s, min_col=5, min_row=25, max_row=27))
    ch2.roundedCorners = False
    ch2.title.overlay = False
    ch2.legend.overlay = False
    ch2.legend.position = "b"
    ch2.x_axis.delete = False
    ch2.y_axis.delete = False
    ch2.height, ch2.width = 8.5, 17
    s.add_chart(ch2, "E29")
    for col_ in "EFGHIJKLMN":
        s.column_dimensions[col_].width = 13

    wb.save(path)


if __name__ == "__main__":
    build(sys.argv[1])
