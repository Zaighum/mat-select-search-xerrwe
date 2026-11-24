"""Generates the Power BI project (PBIP: TMDL semantic model + PBIR report).

Run:  python build_pbip.py <output_dir> <base_theme.json>
"""
import hashlib
import json
import shutil
import sys
import uuid
from pathlib import Path

NAME = "PortOfNEOM"
NS = uuid.UUID("6f1d2c1e-0b7a-4c55-9f0e-0a7e5f3b9c11")
SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition"
PLATFORM = "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json"
NAVY, TEAL = "#0B2545", "#13A89E"


def tag(*parts):
    return str(uuid.uuid5(NS, "/".join(parts)))


def vid(*parts):
    return hashlib.sha1("/".join(parts).encode()).hexdigest()[:20]


# ----------------------------------------------------------------------------
# SEMANTIC MODEL
# ----------------------------------------------------------------------------
# model table -> (excel table, [(column, tmdl type, M type, format, summarizeBy)], [measures])
S, I, D, DT = ("string", "type text"), ("int64", "Int64.Type"), ("double", "type number"), ("dateTime", "type date")
TABLES = {
    "PortProfile": ("tblPortProfile", [
        ("ProfileID", S), ("Category", S), ("Attribute", S), ("Value", S), ("NumericValue", D), ("Unit", S),
        ("SourceID", S), ("Note", S)],
        [("Quay Length (m)", 'CALCULATE(MAX(PortProfile[NumericValue]), PortProfile[ProfileID] = "P04")', "#,##0"),
         ("Berths", 'CALCULATE(MAX(PortProfile[NumericValue]), PortProfile[ProfileID] = "P05")', "0"),
         ("Depth (m)", 'CALCULATE(MAX(PortProfile[NumericValue]), PortProfile[ProfileID] = "P06")', "0.0"),
         ("Saudi Port Volume (TEU)", 'CALCULATE(MAX(PortProfile[NumericValue]), PortProfile[ProfileID] = "P13")', "#,##0"),
         ("Terminal 1 Status", 'CALCULATE(MAX(PortProfile[Value]), PortProfile[ProfileID] = "P03")', None)]),
    "Timeline": ("tblTimeline", [
        ("EventID", S), ("EventDate", DT), ("DatePrecision", S), ("Event", S), ("Category", S), ("Status", S),
        ("SourceID", S)],
        [("Events", "COUNTROWS(Timeline)", "0")]),
    "CapacityPlan": ("tblCapacityPlan", [
        ("CapacityID", S), ("StatementYear", I), ("Milestone", S), ("TargetYear", I), ("CapacityTEU", D),
        ("PlanStatus", S), ("SourceID", S), ("Note", S), ("ShareOfSaudiTEU", D)],
        [("Capacity (TEU)", "SUM(CapacityPlan[CapacityTEU])", "#,##0"),
         ("T1 Phase 1 Capacity (TEU)", 'CALCULATE(SUM(CapacityPlan[CapacityTEU]), CapacityPlan[CapacityID] = "C03")', "#,##0"),
         ("T1 Target Capacity (TEU)", 'CALCULATE(SUM(CapacityPlan[CapacityTEU]), CapacityPlan[CapacityID] = "C04")', "#,##0"),
         ("Phase 1 % of Target", "DIVIDE([T1 Phase 1 Capacity (TEU)], [T1 Target Capacity (TEU)])", "0.0%"),
         ("Phase 1 % of Saudi Volume", "DIVIDE([T1 Phase 1 Capacity (TEU)], [Saudi Port Volume (TEU)])", "0.0%")]),
    "Equipment": ("tblEquipment", [
        ("EquipmentID", S), ("Equipment", S), ("TypeCode", S), ("Units", I), ("Supplier", S), ("Technology", S),
        ("SourceID", S)],
        [("Equipment Units", "SUM(Equipment[Units])", "0"),
         ("STS Cranes", 'CALCULATE(SUM(Equipment[Units]), Equipment[TypeCode] = "STS")', "0")]),
    "Contracts": ("tblContracts", [
        ("ContractID", S), ("Counterparty", S), ("Scope", S), ("Currency", S), ("StatedValueM", D), ("SourceID", S),
        ("Note", S), ("ValueSARm", D), ("ValueUSDm", D)],
        [("Contract Value (SAR m)", "SUM(Contracts[ValueSARm])", "#,##0"),
         ("Contract Value (USD m)", "SUM(Contracts[ValueUSDm])", "#,##0")]),
    "Investment": ("tblInvestment", [
        ("InvestmentID", S), ("Statement", S), ("StatementYear", I), ("AmountSARbn", D), ("SourceID", S), ("Note", S),
        ("AmountUSDbn", D)],
        [("Stated Investment (SAR bn)", "MAX(Investment[AmountSARbn])", "0.0"),
         ("Stated Investment (USD bn)", "MAX(Investment[AmountUSDbn])", "0.00")]),
    "VesselCalls": ("tblVesselCalls", [
        ("CallID", S), ("VesselType", S), ("PeriodGroup", S), ("PeriodDetail", S), ("Calls", I), ("Months", I),
        ("StatedAvgPerMonth", D), ("SourceID", S), ("Note", S), ("AvgCallsPerMonth", D)],
        [("Avg Calls per Month", "SUM(VesselCalls[AvgCallsPerMonth])", "0.0"),
         ("RoRo Calls per Month (Pre-war)", 'CALCULATE([Avg Calls per Month], VesselCalls[CallID] = "V01")', "0.0"),
         ("RoRo Calls (Sep 2026)", 'CALCULATE([Avg Calls per Month], VesselCalls[CallID] = "V02")', "0"),
         ("RoRo Change vs Pre-war", "DIVIDE([RoRo Calls (Sep 2026)], [RoRo Calls per Month (Pre-war)]) - 1", "0%")]),
    "Corridor": ("tblCorridor", [
        ("NodeID", S), ("Sequence", I), ("Node", S), ("Country", S), ("Role", S), ("OnwardMode", S), ("Latitude", D),
        ("Longitude", D), ("SourceID", S)],
        []),
    "Sources": ("tblSources", [
        ("SourceID", S), ("Publisher", S), ("Title", S), ("Published", S), ("URL", S), ("Note", S)],
        [("Sources Cited", "COUNTROWS(Sources)", "0")]),
}
NO_SUM = {"StatementYear", "TargetYear", "Sequence", "Latitude", "Longitude", "NumericValue", "ShareOfSaudiTEU",
          "StatedAvgPerMonth", "Months"}


def table_tmdl(name, excel_table, cols, measures):
    o = [f"table {name}", f"\tlineageTag: {tag('table', name)}", ""]
    for m, expr, fmt in measures:
        o.append(f"\tmeasure '{m}' = {expr}")
        if fmt:
            o.append(f"\t\tformatString: {fmt}")
        o += [f"\t\tlineageTag: {tag('measure', name, m)}", ""]
    for c, (ttype, _) in cols:
        o += [f"\tcolumn {c}", f"\t\tdataType: {ttype}"]
        if ttype == "dateTime":
            o.append("\t\tformatString: dd MMM yyyy")
        elif ttype == "int64":
            o.append("\t\tformatString: 0")
        if c == "URL":
            o.append("\t\tdataCategory: WebUrl")
        if c == "Latitude":
            o.append("\t\tdataCategory: Latitude")
        if c == "Longitude":
            o.append("\t\tdataCategory: Longitude")
        if c == "SourceID" and name == "Sources":
            o.append("\t\tisKey")
        agg = "sum" if ttype in ("int64", "double") and c not in NO_SUM else "none"
        o += [f"\t\tlineageTag: {tag('column', name, c)}", f"\t\tsummarizeBy: {agg}", f"\t\tsourceColumn: {c}", "",
              "\t\tannotation SummarizationSetBy = Automatic", ""]
    types = ", ".join(f'{{"{c}", {mt}}}' for c, (_, mt) in cols)
    names = ", ".join(f'"{c}"' for c, _ in cols)
    o += [f"\tpartition {name} = m", "\t\tmode: import", "\t\tsource =",
          "\t\t\t\tlet",
          "\t\t\t\t    Source = Excel.Workbook(File.Contents(ExcelFilePath), null, true),",
          f'\t\t\t\t    Data = Source{{[Item="{excel_table}",Kind="Table"]}}[Data],',
          f"\t\t\t\t    Selected = Table.SelectColumns(Data, {{{names}}}),",
          f'\t\t\t\t    Typed = Table.TransformColumnTypes(Selected, {{{types}}}, "en-US")',
          "\t\t\t\tin",
          "\t\t\t\t    Typed", "",
          "\tannotation PBI_ResultType = Table", ""]
    return "\n".join(o)


def write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def wjson(path, obj):
    write(path, json.dumps(obj, indent=2, ensure_ascii=False) + "\n")


def build_model(root):
    sm = root / f"{NAME}.SemanticModel"
    wjson(sm / "definition.pbism", {"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json", "version": "4.0", "settings": {}})
    wjson(sm / ".platform", {"$schema": PLATFORM, "metadata": {"type": "SemanticModel", "displayName": NAME},
                             "config": {"version": "2.0", "logicalId": tag("sm")}})
    d = sm / "definition"
    write(d / "database.tmdl", "database\n\tcompatibilityLevel: 1567\n\n")
    order = json.dumps(["ExcelFilePath"] + list(TABLES), separators=(",", ":"))
    model = ["model Model", "\tculture: en-US", "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
             "\tsourceQueryCulture: en-US", "\tdataAccessOptions", "\t\tlegacyRedirects",
             "\t\treturnErrorValuesAsNull", "", "annotation __PBI_TimeIntelligenceEnabled = 0", "",
             f"annotation PBI_QueryOrder = {order}", ""]
    model += [f"ref table {t}" for t in TABLES] + [""]
    write(d / "model.tmdl", "\n".join(model))
    write(d / "expressions.tmdl", "\n".join([
        'expression ExcelFilePath = "C:\\PortOfNEOM\\Port_of_NEOM_Data.xlsx" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]',
        f"\tlineageTag: {tag('expr', 'ExcelFilePath')}", "", "\tannotation PBI_ResultType = Text", "", ""]))
    rel = []
    for t in TABLES:
        if t != "Sources":
            rel += [f"relationship {tag('rel', t)}", f"\tfromColumn: {t}.SourceID", "\ttoColumn: Sources.SourceID", ""]
    write(d / "relationships.tmdl", "\n".join(rel))
    for t, (xl, cols, measures) in TABLES.items():
        write(d / "tables" / f"{t}.tmdl", table_tmdl(t, xl, cols, measures))


# ----------------------------------------------------------------------------
# REPORT
# ----------------------------------------------------------------------------
def lit(v):
    return {"expr": {"Literal": {"Value": v}}}


def colf(t, c):
    return {"Column": {"Expression": {"SourceRef": {"Entity": t}}, "Property": c}}


def meaf(t, m):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": t}}, "Property": m}}


def proj(field, name=None):
    kind = "Column" if "Column" in field else "Measure"
    t = field[kind]["Expression"]["SourceRef"]["Entity"]
    p = {"field": field, "queryRef": f"{t}.{field[kind]['Property']}"}
    if name:
        p["displayName"] = name
    return p


def vtitle(text):
    return {"title": [{"properties": {"show": lit("true"), "text": lit("'" + text.replace("'", "''") + "'")}}]}


def textbox(text, size, color=NAVY, bold=False):
    style = {"fontSize": f"{size}pt", "color": color}
    if bold:
        style["fontWeight"] = "bold"
    return {"visualType": "textbox",
            "objects": {"general": [{"properties": {"paragraphs": [
                {"textRuns": [{"value": text, "textStyle": style}]}]}}]},
            "visualContainerObjects": {"background": [{"properties": {"show": lit("false")}}]},
            "drillFilterOtherVisuals": True}


def card(table, measure, label):
    return {"visualType": "card",
            "query": {"queryState": {"Values": {"projections": [proj(meaf(table, measure), label)]}}},
            "objects": {"labels": [{"properties": {"fontSize": lit("24D")}}]},
            "drillFilterOtherVisuals": True}


def chart(vtype, cat, ys, title_, series=None, sort_by=None):
    qs = {"Category": {"projections": [dict(proj(cat), active=True)]},
          "Y": {"projections": [proj(y) for y in ys]}}
    if series:
        qs["Series"] = {"projections": [proj(series)]}
    q = {"queryState": qs}
    if sort_by:
        q["sortDefinition"] = {"sort": [{"field": sort_by[0], "direction": sort_by[1]}]}
    return {"visualType": vtype, "query": q,
            "objects": {"labels": [{"properties": {"show": lit("true")}}]},
            "visualContainerObjects": vtitle(title_), "drillFilterOtherVisuals": True}


def table(fields, title_, sort_by=None):
    q = {"queryState": {"Values": {"projections": [proj(f, n) for f, n in fields]}}}
    if sort_by:
        q["sortDefinition"] = {"sort": [{"field": sort_by[0], "direction": sort_by[1]}]}
    return {"visualType": "tableEx", "query": q, "visualContainerObjects": vtitle(title_),
            "drillFilterOtherVisuals": True}


def slicer(field, title_):
    return {"visualType": "slicer",
            "query": {"queryState": {"Values": {"projections": [dict(proj(field), active=True)]}}},
            "objects": {"data": [{"properties": {"mode": lit("'Dropdown'")}}]},
            "visualContainerObjects": vtitle(title_), "drillFilterOtherVisuals": True}


def header(page_title):
    return [
        ((24, 12, 900, 44), textbox(page_title, 20, NAVY, True)),
        ((24, 54, 1232, 30), textbox(
            "Public-source data compiled 24 Nov 2026. Capacity figures are plans or operator claims, not measured throughput.",
            10, "#595959")),
    ]


def cards(items, y=92, h=96):
    w = (1232 - 12 * (len(items) - 1)) / len(items)
    return [((24 + i * (w + 12), y, w, h), card(t, m, lbl)) for i, (t, m, lbl) in enumerate(items)]


C, M = colf, meaf
PAGES = [
    ("Overview", header("Port of NEOM - Overview") + cards([
        ("CapacityPlan", "T1 Phase 1 Capacity (TEU)", "Terminal 1 Phase 1 capacity (TEU/yr)"),
        ("CapacityPlan", "T1 Target Capacity (TEU)", "Terminal 1 target capacity, 2030 (TEU/yr)"),
        ("PortProfile", "Quay Length (m)", "Quay wall (m)"),
        ("PortProfile", "Depth (m)", "Draft / basin depth (m)"),
        ("Equipment", "STS Cranes", "Ship-to-shore cranes"),
        ("Investment", "Stated Investment (SAR bn)", "Highest stated investment (SAR bn)"),
    ]) + [
        ((24, 200, 1232, 56), card("PortProfile", "Terminal 1 Status", "Terminal 1 status")),
        ((24, 268, 560, 428), chart("columnChart", C("CapacityPlan", "Milestone"), [M("CapacityPlan", "Capacity (TEU)")],
                                    "Container capacity plan by milestone (TEU per year)",
                                    series=C("CapacityPlan", "PlanStatus"))),
        ((596, 268, 660, 428), table([
            (C("PortProfile", "Category"), "Category"), (C("PortProfile", "Attribute"), "Attribute"),
            (C("PortProfile", "Value"), "Value"), (C("PortProfile", "SourceID"), "Source")],
            "Port profile - key facts")),
    ]),
    ("Infrastructure and Investment", header("Infrastructure and Investment") + cards([
        ("Equipment", "Equipment Units", "Cranes reported ordered (units)"),
        ("Contracts", "Contract Value (SAR m)", "Reported contracts (SAR m, stated minimum)"),
        ("Investment", "Stated Investment (SAR bn)", "Highest stated investment (SAR bn)"),
        ("Investment", "Stated Investment (USD bn)", "Highest stated investment (USD bn)"),
    ]) + [
        ((24, 200, 610, 250), chart("barChart", C("Equipment", "Equipment"), [M("Equipment", "Equipment Units")],
                                    "Cargo-handling equipment by type (units)",
                                    sort_by=(M("Equipment", "Equipment Units"), "Descending"))),
        ((646, 200, 610, 250), chart("columnChart", C("Investment", "Statement"),
                                     [M("Investment", "Stated Investment (SAR bn)")],
                                     "Cumulative investment as stated (SAR bn) - different statements, not a time series",
                                     sort_by=(M("Investment", "Stated Investment (SAR bn)"), "Ascending"))),
        ((24, 462, 1232, 234), table([
            (C("Contracts", "Counterparty"), "Counterparty"), (C("Contracts", "Scope"), "Scope"),
            (M("Contracts", "Contract Value (SAR m)"), "SAR m"), (M("Contracts", "Contract Value (USD m)"), "USD m"),
            (C("Contracts", "Note"), "Note"), (C("Contracts", "SourceID"), "Source")],
            "Reported contracts")),
    ]),
    ("Traffic and Trade Corridor", header("Traffic and Trade Corridor") + cards([
        ("VesselCalls", "RoRo Calls per Month (Pre-war)", "RoRo calls per month, pre-war"),
        ("VesselCalls", "RoRo Calls (Sep 2026)", "RoRo calls, September 2026"),
        ("VesselCalls", "RoRo Change vs Pre-war", "RoRo change vs pre-war"),
        ("CapacityPlan", "Phase 1 % of Saudi Volume", "T1 Phase 1 as % of Saudi port TEU"),
    ]) + [
        ((24, 200, 610, 250), chart("clusteredColumnChart", C("VesselCalls", "VesselType"),
                                    [M("VesselCalls", "Avg Calls per Month")],
                                    "Average vessel calls per month, before and after the war began (Windward)",
                                    series=C("VesselCalls", "PeriodGroup"))),
        ((646, 200, 610, 250), table([
            (C("VesselCalls", "VesselType"), "Vessel type"), (C("VesselCalls", "PeriodDetail"), "Period"),
            (M("VesselCalls", "Avg Calls per Month"), "Calls / month"), (C("VesselCalls", "Note"), "Note")],
            "Vessel calls - detail")),
        ((24, 462, 1232, 234), table([
            (C("Corridor", "Sequence"), "Step"), (C("Corridor", "Node"), "Node"), (C("Corridor", "Country"), "Country"),
            (C("Corridor", "Role"), "Role"), (C("Corridor", "OnwardMode"), "Onward mode")],
            "Europe - Egypt - Port of NEOM - GCC / Iraq land bridge",
            sort_by=(C("Corridor", "Sequence"), "Ascending"))),
    ]),
    ("Timeline and Sources", header("Timeline and Sources") + [
        ((24, 92, 300, 70), slicer(C("Timeline", "Category"), "Category")),
        ((336, 92, 300, 70), slicer(C("Timeline", "Status"), "Status")),
        ((24, 174, 1232, 280), table([
            (C("Timeline", "EventDate"), "Date"), (C("Timeline", "DatePrecision"), "Precision"),
            (C("Timeline", "Event"), "Event"), (C("Timeline", "Category"), "Category"),
            (C("Timeline", "Status"), "Status"), (C("Timeline", "SourceID"), "Source")],
            "Timeline of key events", sort_by=(C("Timeline", "EventDate"), "Ascending"))),
        ((24, 466, 1232, 230), table([
            (C("Sources", "SourceID"), "ID"), (C("Sources", "Publisher"), "Publisher"), (C("Sources", "Title"), "Title"),
            (C("Sources", "Published"), "Published"), (C("Sources", "URL"), "URL")],
            "Sources", sort_by=(C("Sources", "SourceID"), "Ascending"))),
    ]),
]


def build_report(root, theme):
    rp = root / f"{NAME}.Report"
    wjson(root / f"{NAME}.pbip", {"$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json", "version": "1.0", "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
                                  "settings": {"enableAutoRecovery": True}})
    wjson(rp / "definition.pbir", {"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json", "version": "4.0", "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}})
    wjson(rp / ".platform", {"$schema": PLATFORM, "metadata": {"type": "Report", "displayName": NAME},
                             "config": {"version": "2.0", "logicalId": tag("report")}})
    tname = Path(theme).stem
    dest = rp / "StaticResources" / "SharedResources" / "BaseThemes" / Path(theme).name
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(theme, dest)
    d = rp / "definition"
    wjson(d / "version.json", {"$schema": f"{SCHEMA}/versionMetadata/1.0.0/schema.json", "version": "2.0.0"})
    wjson(d / "report.json", {
        "$schema": f"{SCHEMA}/report/1.3.0/schema.json",
        "themeCollection": {"baseTheme": {"name": tname, "reportVersionAtImport": "5.61", "type": "SharedResources"}},
        "layoutOptimization": "None",
        "resourcePackages": [{"name": "SharedResources", "type": "SharedResources",
                              "items": [{"name": tname, "path": f"BaseThemes/{tname}.json", "type": "BaseTheme"}]}],
        "settings": {"useStylableVisualContainerHeader": True, "exportDataMode": "AllowSummarized",
                     "defaultFilterActionIsDataFilter": True, "defaultDrillFilterOtherVisuals": True,
                     "allowChangeFilterTypes": True, "useEnhancedTooltips": True},
    })
    ids = []
    for pname, visuals in PAGES:
        pid = vid("page", pname)
        ids.append(pid)
        wjson(d / "pages" / pid / "page.json", {
            "$schema": f"{SCHEMA}/page/1.4.0/schema.json", "name": pid, "displayName": pname,
            "displayOption": "FitToPage", "height": 720, "width": 1280})
        for i, ((x, y, w, h), visual) in enumerate(visuals):
            v = vid("visual", pname, str(i))
            wjson(d / "pages" / pid / "visuals" / v / "visual.json", {
                "$schema": f"{SCHEMA}/visualContainer/2.0.0/schema.json", "name": v,
                "position": {"x": x, "y": y, "z": i * 1000, "height": h, "width": w, "tabOrder": i * 1000},
                "visual": visual})
    wjson(d / "pages" / "pages.json", {"$schema": f"{SCHEMA}/pagesMetadata/1.0.0/schema.json",
                                       "pageOrder": ids, "activePageName": ids[0]})


if __name__ == "__main__":
    out = Path(sys.argv[1])
    for sub in (f"{NAME}.SemanticModel", f"{NAME}.Report"):
        shutil.rmtree(out / sub, ignore_errors=True)
    build_model(out)
    build_report(out, sys.argv[2])
    write(out / ".gitignore", "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n")
