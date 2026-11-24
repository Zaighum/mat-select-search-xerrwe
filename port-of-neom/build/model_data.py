"""Illustrative commercial / operations / finance model for the Port of NEOM workbook.

NOTHING generated here is an actual Port of NEOM figure. Volumes are synthetic (seeded),
shaped to match the public picture (general/project cargo port, Safaga RoRo shuttle,
2026 bulk surge, near-zero container calls before Terminal 1). Revenue, cost and forecast
cells are Excel formulas driven by the Assumptions sheet.
"""
import math
import random
from datetime import date

from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill

FONT = "Arial"
R0 = 5  # first data row of every table
MONTHS = [date(y, m, 1) for y in range(2023, 2027) for m in range(1, 13) if (y, m) <= (2026, 10)]
N = len(MONTHS)
CARGO = ["General cargo", "Project cargo", "Dry bulk", "RoRo", "Container"]

# cargo type: billing unit, tonnes per unit, handling SAR/unit, storage SAR/unit, marine SAR/call, storage share
CARGO_TYPES = [
    ("General cargo", "tonne", 1.0, 38.0, 4.0, 60000, 0.35),
    ("Project cargo", "freight tonne", 1.0, 55.0, 9.0, 70000, 0.50),
    ("Dry bulk", "tonne", 1.0, 14.0, 1.5, 85000, 0.20),
    ("RoRo", "unit", 20.0, 650.0, 40.0, 28000, 0.10),
    ("Container", "TEU", 12.0, 420.0, 35.0, 95000, 0.60),
]
CT = {c[0]: c for c in CARGO_TYPES}

# id, name (fictional), segment, cargo type, region, contract type, volume share, discount
CUSTOMERS = [
    ("CU01", "Construction Logistics Co. A", "Construction supply", "General cargo", "Saudi Arabia", "Annual volume agreement", 0.40, 0.08),
    ("CU02", "Steel & Materials Importer B", "Industrial importer", "General cargo", "East Asia", "Annual volume agreement", 0.35, 0.06),
    ("CU03", "Regional Freight Forwarder C", "Freight forwarder", "General cargo", "GCC", "Public tariff", 0.25, 0.00),
    ("CU04", "Heavy-Lift Project Forwarder D", "Project forwarder", "Project cargo", "Europe", "Project contract", 0.55, 0.05),
    ("CU05", "Energy EPC Contractor E", "EPC contractor", "Project cargo", "East Asia", "Project contract", 0.45, 0.04),
    ("CU06", "Cement & Aggregates Trader F", "Bulk trader", "Dry bulk", "Egypt", "Spot", 0.60, 0.00),
    ("CU07", "Grain Importer G", "Bulk trader", "Dry bulk", "GCC", "Spot", 0.40, 0.00),
    ("CU08", "Red Sea Ferry Operator H", "RoRo operator", "RoRo", "Egypt", "Liner service agreement", 0.70, 0.10),
    ("CU09", "Regional Trucking Consortium I", "Road haulier", "RoRo", "GCC", "Public tariff", 0.30, 0.00),
    ("CU10", "Regional Feeder Line J", "Shipping line", "Container", "Red Sea", "Liner service agreement", 0.65, 0.07),
    ("CU11", "Feeder Line K", "Shipping line", "Container", "Red Sea", "Public tariff", 0.35, 0.00),
]

ACCOUNTS = [
    ("R1", "Cargo handling revenue", "Revenue", 1), ("R2", "Marine services revenue", "Revenue", 2),
    ("R3", "Storage and ancillary revenue", "Revenue", 3), ("R4", "Lease and other income", "Revenue", 4),
    ("O1", "Staff costs", "Opex", 5), ("O2", "Equipment maintenance", "Opex", 6), ("O3", "Energy and fuel", "Opex", 7),
    ("O4", "Outsourced labour", "Opex", 8), ("O5", "Marine operations", "Opex", 9),
    ("O6", "General and administration", "Opex", 10), ("D1", "Depreciation", "Depreciation", 11),
]

# key, label, value, number format, note
ASSUMPTIONS = [
    ("ESC", "Tariff escalation per year", 0.03, "0.0%", "Applied to all rates from 2023."),
    ("INFL", "Cost inflation per year", 0.03, "0.0%", "Applied to fixed and unit costs from 2023."),
    ("STAFF", "Staff cost per head per month (SAR)", 14500, "#,##0", "Fully loaded."),
    ("MAINT_FIX", "Maintenance - fixed per month (SAR)", 2200000, "#,##0", ""),
    ("MAINT_VAR", "Maintenance - per tonne handled (SAR)", 1.1, "0.00", ""),
    ("ENERGY", "Energy and fuel per tonne (SAR)", 2.4, "0.00", ""),
    ("LABOUR", "Outsourced labour per tonne (SAR)", 6.5, "0.00", ""),
    ("MARINE_OPS", "Marine operations cost per vessel call (SAR)", 9000, "#,##0", "Tugs, pilot boats, mooring."),
    ("GA", "General and administration per month (SAR)", 3800000, "#,##0", ""),
    ("LEASE", "Lease and other income per month, 2023 (SAR)", 1500000, "#,##0", ""),
    ("LEASE_G", "Lease income growth per year", 0.05, "0.0%", ""),
    ("BERTHS", "General-purpose berths in use", 5, "0", "Denominator for berth occupancy."),
    ("HOURS_HEAD", "Worked hours per head per month", 190, "0", "For the injury frequency rate."),
    ("ASSET_LIFE", "Average asset life (years)", 30, "0", "Straight-line depreciation."),
    ("ASSET_2023", "Assets in service 2023 (SAR)", 2500000000, "#,##0", "Terminal 1 excluded until it opens."),
    ("ASSET_2024", "Assets in service 2024 (SAR)", 3400000000, "#,##0", ""),
    ("ASSET_2025", "Assets in service 2025 (SAR)", 4200000000, "#,##0", ""),
    ("ASSET_2026", "Assets in service 2026 (SAR)", 4800000000, "#,##0", ""),
    ("MONTHS_2026", "Months of 2026 actuals in the data", 10, "0", "Used to annualise 2026 for the forecast base."),
    ("YIELD_TEU", "Forecast: all-in container yield per TEU, 2026 prices (SAR)", 560, "#,##0", "Handling + storage + marine."),
    ("T1_FIXED", "Forecast: Terminal 1 fixed opex per year (SAR)", 160000000, "#,##0", "Once operating."),
    ("T1_VAR", "Forecast: Terminal 1 variable cost per TEU (SAR)", 110, "#,##0", ""),
    ("FIXED_SHARE", "Forecast: share of existing opex that is fixed", 0.70, "0%", "Remainder scales with tonnage."),
]

# scenario -> (utilisation of T1 capacity, non-container tonnage growth, capex SAR m) for 2027..2031
FC_YEARS = [2027, 2028, 2029, 2030, 2031]
FC_CAPACITY = [700000, 700000, 700000, 1500000, 1500000]  # public plan: 700k phase 1, 1.5m by 2030
SCENARIOS = {
    "Base": ([0.15, 0.30, 0.45, 0.35, 0.45], [-0.08, 0.02, 0.02, 0.02, 0.02], [900, 700, 800, 500, 300]),
    "High": ([0.25, 0.45, 0.65, 0.50, 0.62], [0.00, 0.05, 0.05, 0.05, 0.05], [900, 800, 1000, 600, 400]),
    "Low": ([0.08, 0.18, 0.28, 0.20, 0.27], [-0.15, -0.03, -0.03, 0.00, 0.00], [900, 600, 500, 300, 200]),
}

CAPEX_PROJECTS = [  # name, total budget SAR, first month index, last month index
    ("Terminal 1 quay and civil works", 3.2e9, 0, 40),
    ("Terminal 1 cranes and equipment", 1.6e9, 14, 45),
    ("Dredging and access channel", 0.9e9, 0, 22),
    ("General cargo berths and yard upgrade", 0.7e9, 2, 34),
    ("Digital and automation systems", 0.4e9, 12, 45),
]


# ----------------------------------------------------------------------------
# synthetic volumes
# ----------------------------------------------------------------------------
def lerp(i, pts):
    for (i0, v0), (i1, v1) in zip(pts, pts[1:]):
        if i0 <= i <= i1:
            return v0 + (v1 - v0) * (i - i0) / (i1 - i0)
    return pts[-1][1]


RORO_CALLS = ([10, 10, 11, 11, 12, 12, 12, 13, 13, 13, 14, 14] + [14, 14, 15, 14, 15, 15, 16, 15, 16, 16, 16, 16]
              + [17, 16, 18, 17, 17, 16, 18, 17, 17, 18, 16, 17] + [17, 17, 20, 23, 25, 26, 27, 28, 30, 31])
BULK_CALLS = ([1, 1, 1, 2, 1, 1, 1, 1, 2, 1, 1, 1] + [1, 1, 2, 1, 1, 1, 1, 1, 1, 2, 1, 1]
              + [0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0] + [0, 0, 4, 5, 5, 5, 5, 4, 4, 3])
CONT_CALLS = ([0, 1, 1, 1, 1, 1, 1, 1, 1, 2, 1, 1] + [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2]
              + [1, 2, 1, 1, 2, 1, 2, 1, 1, 1, 2, 1] + [2, 1, 1, 0, 1, 0, 0, 0, 0, 0])
WAR_START = MONTHS.index(date(2026, 3, 1))


def volumes(actual=True):
    """Returns {cargo: [(volume, calls, berth_hours)] per month}. actual=False gives the budget view
    (smooth trend, no 2026 disruption)."""
    rnd = random.Random(20261124)
    nz = (lambda s: 1 + rnd.uniform(-s, s)) if actual else (lambda s: 1.0)
    out = {c: [] for c in CARGO}
    for i, d in enumerate(MONTHS):
        season = 1 + 0.05 * math.sin(2 * math.pi * (d.month - 1) / 12)
        g = lerp(i, [(0, 180000), (23, 320000), (35, 300000), (45, 265000)]) * season * nz(0.06)
        gc = max(1, round(g / 9000))
        out["General cargo"].append((round(g, -2), gc, round(gc * 36 * nz(0.10))))
        p = lerp(i, [(0, 60000), (20, 140000), (45, 90000)]) * season * nz(0.09)
        pc = max(1, round(p / 6000))
        out["Project cargo"].append((round(p, -2), pc, round(pc * 48 * nz(0.12))))
        if actual:
            bc, rc, cc = BULK_CALLS[i], RORO_CALLS[i], CONT_CALLS[i]
            b_hours = 190 if i >= WAR_START else 60
        else:
            bc = 1
            rc = round(lerp(i, [(0, 10), (24, 16), (45, 19)]))
            cc = 2 if i % 2 else 1
            b_hours = 60
        out["Dry bulk"].append((round(bc * 38000 * nz(0.10), -2), bc, round(bc * b_hours * nz(0.15))))
        out["RoRo"].append((round(rc * 76 * nz(0.05)), rc, round(rc * 9 * nz(0.10))))
        out["Container"].append((round(cc * 420 * nz(0.15)), cc, round(cc * 16 * nz(0.15))))
    return out


def kpi_inputs(actual=True):
    rnd = random.Random(7)
    rows = []
    for i in range(N):
        surge = i >= WAR_START
        head = 420 + 8 * i
        if not actual:
            rows.append((head,))
            continue
        rows.append((head,
                     round((14 + (i - WAR_START) * 1.1 if surge else 6) + rnd.uniform(-2, 2), 1),
                     round(17 + 0.12 * i + rnd.uniform(-1, 1), 1),
                     round(48 - 0.35 * i + (8 if surge else 0) + rnd.uniform(-2, 2), 1),
                     round(0.90 + rnd.uniform(-0.03, 0.04), 3),
                     rnd.choices([0, 1, 2], [0.80, 0.18, 0.02])[0]))
    return rows


def py_finance(vol, heads, a):
    """Python mirror of the Excel finance formulas; used for the budget and to cross-check the recalc."""
    disc = {c: sum(cu[6] * cu[7] for cu in CUSTOMERS if cu[3] == c) for c in CARGO}
    res = []
    for i, d in enumerate(MONTHS):
        k = d.year - 2023
        esc, infl = (1 + a["ESC"]) ** k, (1 + a["INFL"]) ** k
        r1 = r2 = r3 = tonnes = calls = 0.0
        for c in CARGO:
            v, n, _ = vol[c][i]
            _, _, tpu, hand, stor, marine, share = CT[c]
            r1 += v * hand * esc * (1 - disc[c])
            r2 += n * marine * esc
            r3 += v * share * stor * esc
            tonnes += v * tpu
            calls += n
        res.append({
            "R1": r1, "R2": r2, "R3": r3, "R4": a["LEASE"] * (1 + a["LEASE_G"]) ** k,
            "O1": heads[i] * a["STAFF"] * infl, "O2": (a["MAINT_FIX"] + a["MAINT_VAR"] * tonnes) * infl,
            "O3": a["ENERGY"] * tonnes * infl, "O4": a["LABOUR"] * tonnes * infl,
            "O5": a["MARINE_OPS"] * calls * infl, "O6": a["GA"] * infl,
            "D1": a[f"ASSET_{d.year}"] / a["ASSET_LIFE"] / 12,
        })
    return res


def capex_rows():
    rnd = random.Random(99)
    rows, k = [], 0
    for name, total, i0, i1 in CAPEX_PROJECTS:
        w = [math.sin(math.pi * (j + 0.5) / (i1 - i0 + 1)) for j in range(i1 - i0 + 1)]
        bud = {i0 + j: total * x / sum(w) for j, x in enumerate(w)}
        for i in range(N):
            b = bud.get(i, 0.0)
            act = bud.get(i - 2, 0.0) * 1.06 * (1 + rnd.uniform(-0.12, 0.12))  # 2-month slippage, 6% overrun
            if b or act:
                k += 1
                rows.append((f"X{k:03d}", MONTHS[i], name, round(b, -3), round(act, -3)))
    return rows


# ----------------------------------------------------------------------------
# workbook
# ----------------------------------------------------------------------------
def rg(sheet, col, n):
    return f"{sheet}!${col}${R0}:${col}${R0 + n - 1}"


def add_model(wb, data_sheet, wa):
    """Adds the illustrative sheets. wa = Assumptions sheet. Returns the assumption values."""
    WARN = "ILLUSTRATIVE - generated data, not actual Port of NEOM figures. "
    orange = "ED7D31"

    # ---- assumptions -------------------------------------------------------
    wa["A6"] = "Illustrative model drivers (blue = input)"
    wa["A6"].font = Font(name=FONT, bold=True, color="0B2545")
    A, vals = {}, {}
    for r, (key, label, val, fmt, note) in enumerate(ASSUMPTIONS, 7):
        wa.cell(r, 1, label).font = Font(name=FONT, size=10)
        c = wa.cell(r, 2, val)
        c.font = Font(name=FONT, color="0000FF", bold=True)
        c.number_format = fmt
        wa.cell(r, 3, note).font = Font(name=FONT, size=10, italic=True)
        A[key], vals[key] = f"Assumptions!$B${r}", val
    wa.column_dimensions["A"].width = 58
    wa.column_dimensions["B"].width = 18

    # ---- dimensions --------------------------------------------------------
    cal = [(d, d.year, f"Q{(d.month - 1) // 3 + 1}", d.month, d.strftime("%b"), d.strftime("%Y-%m")) for d in MONTHS]
    data_sheet(wb, "Calendar", "tblCalendar", ["Date", "Year", "Quarter", "MonthNo", "MonthName", "YearMonth"], cal,
               [13, 8, 9, 9, 11, 11], "Month calendar for the illustrative model (Jan 2023 - Oct 2026).",
               {"Date": "dd-mmm-yyyy"})
    data_sheet(wb, "CargoTypes", "tblCargoTypes",
               ["CargoType", "BillingUnit", "TonnesPerUnit", "HandlingRateSAR", "StorageRateSAR", "MarineRatePerCallSAR", "StorageShare"],
               CARGO_TYPES, [16, 14, 14, 16, 15, 20, 13],
               WARN + "Tariff rates are assumptions at 2023 prices, not the published Port of NEOM tariff.",
               {"StorageShare": "0%", "MarineRatePerCallSAR": "#,##0"})
    data_sheet(wb, "Customers", "tblCustomers",
               ["CustomerID", "Customer", "Segment", "CargoType", "Region", "ContractType", "VolumeShare", "DiscountPct"],
               CUSTOMERS, [11, 32, 20, 15, 14, 24, 12, 12],
               WARN + "Customers are fictional placeholders. VolumeShare sums to 100% within each cargo type.",
               {"VolumeShare": "0%", "DiscountPct": "0%"})
    data_sheet(wb, "Accounts", "tblAccounts", ["AccountID", "Account", "Category", "SortOrder"], ACCOUNTS,
               [10, 32, 14, 10], "Chart of accounts for the illustrative P&L.")

    # ---- operations --------------------------------------------------------
    vol = volumes(True)
    n_ops = N * len(CARGO)
    rows, k = [], 0
    for i, d in enumerate(MONTHS):
        for c in CARGO:
            k += 1
            r = R0 + k - 1
            v, n, bh = vol[c][i]
            rows.append((f"OP{k:04d}", d, c, v, n, bh,
                         f"=D{r}*INDEX({rg('CargoTypes', 'C', 5)},MATCH(C{r},{rg('CargoTypes', 'A', 5)},0))"))
    data_sheet(wb, "Operations", "tblOperations",
               ["OpsID", "Month", "CargoType", "Volume", "VesselCalls", "BerthHours", "Tonnes"], rows,
               [9, 13, 15, 12, 11, 11, 12],
               WARN + "Volume is in the cargo type's billing unit (see CargoTypes); Tonnes converts it.",
               {"Month": "mmm-yyyy", "Volume": "#,##0", "Tonnes": "#,##0", "BerthHours": "#,##0"})
    OPS = {c: rg("Operations", c, n_ops) for c in "BCDEFG"}

    kin = kpi_inputs(True)
    rows = []
    for i, d in enumerate(MONTHS):
        r = R0 + i
        rows.append((d,) + kin[i] + (
            f"=SUMIFS({OPS['F']},{OPS['B']},A{r})", f"=SUMIFS({OPS['E']},{OPS['B']},A{r})",
            f"=H{r}/({A['BERTHS']}*DAY(EOMONTH(A{r},0))*24)", f"=B{r}*{A['HOURS_HEAD']}"))
    data_sheet(wb, "OpsKPI", "tblOpsKPI",
               ["Month", "Headcount", "AnchorageWaitHrs", "CraneMovesPerHour", "TruckTurnMin", "EquipmentAvailability",
                "LostTimeInjuries", "BerthHours", "VesselCalls", "BerthOccupancy", "ManHours"], rows,
               [13, 11, 16, 17, 13, 20, 15, 11, 11, 14, 11],
               WARN + "Monthly port-level KPIs. Last four columns are formulas over Operations and Assumptions.",
               {"Month": "mmm-yyyy", "EquipmentAvailability": "0.0%", "BerthOccupancy": "0.0%", "BerthHours": "#,##0",
                "ManHours": "#,##0"})

    # ---- commercial --------------------------------------------------------
    CTc = {c: rg("CargoTypes", c, 5) for c in "ACDEFG"}
    CUc = {c: rg("Customers", c, len(CUSTOMERS)) for c in "AGH"}
    rows, k = [], 0
    for d in MONTHS:
        for cu in CUSTOMERS:
            k += 1
            r = R0 + k - 1
            ct = lambda col: f"INDEX({CTc[col]},MATCH($D{r},{CTc['A']},0))"
            cuf = lambda col: f"INDEX({CUc[col]},MATCH($C{r},{CUc['A']},0))"
            esc = f"(1+{A['ESC']})^(YEAR($B{r})-2023)"
            rows.append((
                f"RV{k:04d}", d, cu[0], cu[3],
                f"={cuf('G')}*SUMIFS({OPS['D']},{OPS['B']},$B{r},{OPS['C']},$D{r})",
                f"={cuf('G')}*SUMIFS({OPS['E']},{OPS['B']},$B{r},{OPS['C']},$D{r})",
                f"=E{r}*{ct('C')}",
                f"=E{r}*{ct('D')}*{esc}",
                f"=H{r}*{cuf('H')}",
                f"=F{r}*{ct('F')}*{esc}",
                f"=E{r}*{ct('G')}*{ct('E')}*{esc}",
                f"=H{r}-I{r}+J{r}+K{r}"))
    n_rev = len(rows)
    data_sheet(wb, "Commercial", "tblCommercial",
               ["RevenueID", "Month", "CustomerID", "CargoType", "Volume", "AttributedCalls", "Tonnes", "GrossHandlingSAR",
                "DiscountSAR", "MarineRevenueSAR", "StorageRevenueSAR", "NetRevenueSAR"], rows,
               [10, 13, 11, 15, 12, 14, 12, 17, 13, 17, 17, 15],
               WARN + "Revenue by customer and month = customer share of Operations volume x CargoTypes rates, escalated.",
               {"Month": "mmm-yyyy", **{h: "#,##0" for h in ["Volume", "Tonnes", "GrossHandlingSAR", "DiscountSAR",
                                                             "MarineRevenueSAR", "StorageRevenueSAR", "NetRevenueSAR"]},
                "AttributedCalls": "0.0"})
    COM = {c: rg("Commercial", c, n_rev) for c in "BHIJKL"}

    # ---- finance -----------------------------------------------------------
    KPIc = {c: rg("OpsKPI", c, N) for c in "AB"}
    budget = py_finance(volumes(False), [h[0] for h in kpi_inputs(False)], vals)
    rows, k = [], 0
    for i, d in enumerate(MONTHS):
        for acc in ACCOUNTS:
            k += 1
            r = R0 + k - 1
            infl = f"(1+{A['INFL']})^(YEAR(B{r})-2023)"
            ton = f"SUMIFS({OPS['G']},{OPS['B']},B{r})"
            f = {
                "R1": f"=SUMIFS({COM['H']},{COM['B']},B{r})-SUMIFS({COM['I']},{COM['B']},B{r})",
                "R2": f"=SUMIFS({COM['J']},{COM['B']},B{r})",
                "R3": f"=SUMIFS({COM['K']},{COM['B']},B{r})",
                "R4": f"={A['LEASE']}*(1+{A['LEASE_G']})^(YEAR(B{r})-2023)",
                "O1": f"=INDEX({KPIc['B']},MATCH(B{r},{KPIc['A']},0))*{A['STAFF']}*{infl}",
                "O2": f"=({A['MAINT_FIX']}+{A['MAINT_VAR']}*{ton})*{infl}",
                "O3": f"={A['ENERGY']}*{ton}*{infl}",
                "O4": f"={A['LABOUR']}*{ton}*{infl}",
                "O5": f"={A['MARINE_OPS']}*SUMIFS({OPS['E']},{OPS['B']},B{r})*{infl}",
                "O6": f"={A['GA']}*{infl}",
                "D1": f"={A['ASSET_' + str(d.year)]}/{A['ASSET_LIFE']}/12",
            }[acc[0]]
            rows.append((f"F{k:04d}", d, acc[0], "Actual", f))
    n_act = k
    for i, d in enumerate(MONTHS):
        for acc in ACCOUNTS:
            k += 1
            rows.append((f"F{k:04d}", d, acc[0], "Budget", round(budget[i][acc[0]], -2)))
    n_fin = k
    data_sheet(wb, "Finance", "tblFinance", ["FinanceID", "Month", "AccountID", "Scenario", "AmountSAR"], rows,
               [10, 13, 11, 11, 16],
               WARN + "Actual rows are formulas (revenue ties to Commercial; costs use Assumptions drivers). Budget rows are inputs: "
                      "the same model run on a smooth pre-2026 volume plan.",
               {"Month": "mmm-yyyy", "AmountSAR": "#,##0"})
    FIN = {c: rg("Finance", c, n_fin) for c in "BCDE"}

    cx = capex_rows()
    data_sheet(wb, "Capex", "tblCapex", ["CapexID", "Month", "Project", "BudgetSAR", "ActualSAR"], cx,
               [9, 13, 40, 16, 16], WARN + "Monthly capital expenditure by project (inputs).",
               {"Month": "mmm-yyyy", "BudgetSAR": "#,##0", "ActualSAR": "#,##0"})
    CAP = {c: rg("Capex", c, len(cx)) for c in "BDE"}

    # ---- forecast base (formulas on Assumptions) ---------------------------
    y26 = '">="&DATE(2026,1,1)'
    ann = f"*12/{A['MONTHS_2026']}"
    base = [
        ("BASE_TONNES", "Forecast base: 2026 tonnes, annualised", f"=SUMIFS({OPS['G']},{OPS['B']},{y26}){ann}", "#,##0"),
        ("BASE_REV", "Forecast base: 2026 cargo revenue, annualised (SAR)", f"=SUMIFS({COM['L']},{COM['B']},{y26}){ann}", "#,##0"),
        ("BASE_OPEX", "Forecast base: 2026 opex, annualised (SAR)",
         f'=SUMIFS({FIN["E"]},{FIN["D"]},"Actual",{FIN["B"]},{y26},{FIN["C"]},"O*"){ann}', "#,##0"),
    ]
    r = 7 + len(ASSUMPTIONS) + 1
    wa.cell(r - 1, 1, "Forecast base (formulas over 2026 actuals)").font = Font(name=FONT, bold=True, color="0B2545")
    for key, label, f, fmt in base:
        wa.cell(r, 1, label).font = Font(name=FONT, size=10)
        wa.cell(r, 2, f).number_format = fmt
        wa.cell(r, 2).font = Font(name=FONT, bold=True)
        A[key] = f"Assumptions!$B${r}"
        r += 1
    wa.cell(r, 1, "Forecast base: cargo revenue per tonne (SAR)").font = Font(name=FONT, size=10)
    wa.cell(r, 2, f"={A['BASE_REV']}/{A['BASE_TONNES']}").number_format = "0.00"
    wa.cell(r, 2).font = Font(name=FONT, bold=True)
    A["BASE_YIELD"] = f"Assumptions!$B${r}"

    rows, k = [], 0
    for sc, (util, growth, capex) in SCENARIOS.items():
        for j, y in enumerate(FC_YEARS):
            k += 1
            r = R0 + k - 1
            esc, infl = f"(1+{A['ESC']})^(A{r}-2026)", f"(1+{A['INFL']})^(A{r}-2026)"
            prev = A["BASE_TONNES"] if j == 0 else f"H{r - 1}"
            rows.append((
                y, sc, FC_CAPACITY[j], util[j], growth[j], capex[j],
                f"=C{r}*D{r}", f"={prev}*(1+E{r})",
                f"=G{r}*{A['YIELD_TEU']}*{esc}/1000000",
                f"=H{r}*{A['BASE_YIELD']}*{esc}/1000000",
                f"={A['LEASE']}*12*(1+{A['LEASE_G']})^(A{r}-2023)/1000000",
                f"=I{r}+J{r}+K{r}",
                f"=({A['BASE_OPEX']}*({A['FIXED_SHARE']}+(1-{A['FIXED_SHARE']})*H{r}/{A['BASE_TONNES']})"
                f"+IF(G{r}>0,{A['T1_FIXED']}+{A['T1_VAR']}*G{r},0))*{infl}/1000000",
                f"=L{r}-M{r}", f"=N{r}/L{r}"))
    data_sheet(wb, "Forecast", "tblForecast",
               ["Year", "Scenario", "T1CapacityTEU", "T1Utilisation", "NonContainerGrowth", "CapexSARm", "ContainerTEU",
                "NonContainerTonnes", "ContainerRevenueSARm", "NonContainerRevenueSARm", "LeaseIncomeSARm", "RevenueSARm",
                "OpexSARm", "EBITDASARm", "EBITDAMargin"], rows,
               [7, 10, 14, 13, 18, 11, 13, 18, 20, 23, 16, 13, 11, 12, 13],
               WARN + "Five-year forecast, three scenarios. Columns C-F are inputs; the rest are formulas. T1 capacity follows the public plan.",
               {"T1CapacityTEU": "#,##0", "T1Utilisation": "0%", "NonContainerGrowth": "0%", "ContainerTEU": "#,##0",
                "NonContainerTonnes": "#,##0", "EBITDAMargin": "0.0%",
                **{h: "#,##0.0" for h in ["ContainerRevenueSARm", "NonContainerRevenueSARm", "LeaseIncomeSARm",
                                          "RevenueSARm", "OpexSARm", "EBITDASARm"]}})

    for name in ["Calendar", "CargoTypes", "Customers", "Accounts", "Operations", "OpsKPI", "Commercial", "Finance",
                 "Capex", "Forecast"]:
        wb[name].sheet_properties.tabColor = orange
        wb[name]["A2"].font = Font(name=FONT, size=9, italic=True, color="C00000")

    # ---- model summary -----------------------------------------------------
    s = wb.create_sheet("ModelSummary", 2)
    s.sheet_view.showGridLines = False
    s.sheet_properties.tabColor = orange
    s["A1"] = "Illustrative model - yearly summary (SAR million unless stated)"
    s["A1"].font = Font(name=FONT, size=14, bold=True, color="0B2545")
    s["A2"] = WARN + "2026 covers January to October."
    s["A2"].font = Font(name=FONT, size=9, italic=True, color="C00000")
    s.column_dimensions["A"].width = 40
    years = [2023, 2024, 2025, 2026]
    s["A4"] = "Measure"
    for j, y in enumerate(years):
        s.cell(4, 2 + j, y)
        s.column_dimensions[chr(66 + j)].width = 14
    for c in range(1, 6):
        s.cell(4, c).font = Font(name=FONT, bold=True, color="FFFFFF")
        s.cell(4, c).fill = PatternFill("solid", fgColor="0B2545")

    def yr(col):
        return f'">="&DATE({col}$4,1,1),{{B}},"<="&DATE({col}$4,12,31)'

    def fin(col, pat, scen="Actual"):
        return (f'SUMIFS({FIN["E"]},{FIN["D"]},"{scen}",{FIN["C"]},"{pat}",{FIN["B"]},'
                + yr(col).replace("{B}", FIN["B"]) + ")/1000000")

    lines = [
        ("Tonnes handled (thousand)", lambda c: f'=SUMIFS({OPS["G"]},{OPS["B"]},' + yr(c).replace("{B}", OPS["B"]) + ")/1000", "#,##0"),
        ("Vessel calls", lambda c: f'=SUMIFS({OPS["E"]},{OPS["B"]},' + yr(c).replace("{B}", OPS["B"]) + ")", "#,##0"),
        ("Revenue - actual", lambda c: "=" + fin(c, "R*"), "#,##0.0"),
        ("Revenue - budget", lambda c: "=" + fin(c, "R*", "Budget"), "#,##0.0"),
        ("Revenue variance vs budget", lambda c: f"={c}7/{c}8-1", "0.0%"),
        ("Operating costs - actual", lambda c: "=" + fin(c, "O*"), "#,##0.0"),
        ("EBITDA - actual", lambda c: f"={c}7-{c}10", "#,##0.0"),
        ("EBITDA margin", lambda c: f"={c}11/{c}7", "0.0%"),
        ("Depreciation", lambda c: "=" + fin(c, "D*"), "#,##0.0"),
        ("EBIT", lambda c: f"={c}11-{c}13", "#,##0.0"),
        ("Capex - actual", lambda c: f'=SUMIFS({CAP["E"]},{CAP["B"]},' + yr(c).replace("{B}", CAP["B"]) + ")/1000000", "#,##0"),
        ("Capex - budget", lambda c: f'=SUMIFS({CAP["D"]},{CAP["B"]},' + yr(c).replace("{B}", CAP["B"]) + ")/1000000", "#,##0"),
        ("Check: finance cargo revenue less Commercial (should be 0)",
         lambda c: "=" + fin(c, "R1") + "+" + fin(c, "R2") + "+" + fin(c, "R3")
         + f'-SUMIFS({COM["L"]},{COM["B"]},' + yr(c).replace("{B}", COM["B"]) + ")/1000000", "0.000"),
    ]
    for i, (label, f, fmt) in enumerate(lines, 5):
        s.cell(i, 1, label).font = Font(name=FONT, size=10, bold=label.startswith(("EBITDA -", "Revenue - actual")))
        for j in range(4):
            c = s.cell(i, 2 + j, f(chr(66 + j)))
            c.number_format = fmt
            c.font = Font(name=FONT, size=10)
    ch = BarChart()
    ch.type = "col"
    ch.title = "Illustrative revenue and EBITDA by year (SAR million)"
    for row in (7, 11):
        ch.add_data(Reference(s, min_col=1, max_col=5, min_row=row), titles_from_data=True, from_rows=True)
    ch.set_categories(Reference(s, min_col=2, max_col=5, min_row=4))
    ch.roundedCorners = False
    ch.title.overlay = False
    ch.legend.overlay = False
    ch.legend.position = "b"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.height, ch.width = 8.5, 17
    s.add_chart(ch, "G4")
    return vals
