"""Renders PNG previews of the report pages from the same layout (build_pbip.PAGES) and the same
workbook data. These are NOT Power BI screenshots: Power BI Desktop will style things differently.

Run:  python render_preview.py <workbook.xlsx> <out_dir>
"""
import re
import sys
import textwrap
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.ticker
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.patches import FancyBboxPatch
from openpyxl import load_workbook

import build_pbip as b

W, H = 1280, 720
NAVY, GREY = "#0B2545", "#595959"
PALETTE = ["#0B2545", "#13A89E", "#F2A541", "#8D6CAB", "#D1495B", "#6C9BD2", "#7A8B99", "#A3C4BC", "#E8B4B8", "#3D5A80"]
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 7})


# ---------------------------------------------------------------- data
def load(path):
    wb = load_workbook(path, data_only=True)
    raw = {}
    for ws in wb:
        for t in ws.tables.values():
            rows = [[c.value for c in r] for r in ws[t.ref]]
            raw[t.displayName] = pd.DataFrame(rows[1:], columns=rows[0])
    T = {name: raw[xl].copy() for name, (xl, _, _) in b.TABLES.items()}
    F = {}
    for name, df in T.items():
        d = df.copy()
        d.columns = [f"{name}.{c}" for c in d.columns]
        F[name] = d
    for (ft, fc), (tt, tc) in b.RELATIONSHIPS:  # denormalise dimensions onto facts
        dim = T[tt].copy()
        dim.columns = [f"{tt}.{c}" for c in dim.columns]
        F[ft] = F[ft].merge(dim, how="left", left_on=f"{ft}.{fc}", right_on=f"{tt}.{tc}")
    return F


def flt(F, table, col, value):
    key = f"{table}.{col}"
    return {n: (d[d[key] == value] if key in d.columns else d) for n, d in F.items()}


def col(F, t, c):
    return F[t][f"{t}.{c}"]


def ssum(F, t, c, scale=1.0, **eq):
    d = F[t]
    for k, v in eq.items():
        d = d[d[k.replace("__", ".")] == v]
    return None if d.empty else d[f"{t}.{c}"].sum() / scale


def div(a, c):
    return None if a is None or not c else a / c


def sub(a, c):
    return None if a is None and c is None else (a or 0) - (c or 0)


def fin(F, scen, cat=None):
    eq = {"Finance__Scenario": scen}
    if cat:
        eq["Accounts__Category"] = cat
    return ssum(F, "Finance", "AmountSAR", 1e6, **eq)


def prof(F, pid, c="NumericValue"):
    d = F["PortProfile"]
    d = d[d["PortProfile.ProfileID"] == pid]
    return None if d.empty else d[f"PortProfile.{c}"].iloc[0]


def fc(F, c):
    return ssum(F, "Forecast", c)


def base31(F, fn):
    return fn(flt(flt(F, "Forecast", "Scenario", "Base"), "Forecast", "Year", 2031))


def mean(F, t, c):
    return None if F[t].empty else col(F, t, c).mean()


MEASURES = {
    "Revenue (SAR m)": lambda F: fin(F, "Actual", "Revenue"),
    "Budget Revenue (SAR m)": lambda F: fin(F, "Budget", "Revenue"),
    "Opex (SAR m)": lambda F: fin(F, "Actual", "Opex"),
    "Budget Opex (SAR m)": lambda F: fin(F, "Budget", "Opex"),
    "Depreciation (SAR m)": lambda F: fin(F, "Actual", "Depreciation"),
    "Actual (SAR m)": lambda F: fin(F, "Actual"),
    "Budget (SAR m)": lambda F: fin(F, "Budget"),
    "Tonnes (k)": lambda F: ssum(F, "Operations", "Tonnes", 1e3),
    "Vessel Calls": lambda F: ssum(F, "Operations", "VesselCalls"),
    "Berth Hours": lambda F: ssum(F, "Operations", "BerthHours"),
    "Berth Occupancy": lambda F: mean(F, "OpsKPI", "BerthOccupancy"),
    "Crane Moves per Hour": lambda F: mean(F, "OpsKPI", "CraneMovesPerHour"),
    "Truck Turn (min)": lambda F: mean(F, "OpsKPI", "TruckTurnMin"),
    "Anchorage Wait (hrs)": lambda F: mean(F, "OpsKPI", "AnchorageWaitHrs"),
    "LTIFR": lambda F: div(ssum(F, "OpsKPI", "LostTimeInjuries", 1e-6), ssum(F, "OpsKPI", "ManHours")),
    "Net Revenue (SAR m)": lambda F: ssum(F, "Commercial", "NetRevenueSAR", 1e6),
    "Marine Revenue (SAR m)": lambda F: ssum(F, "Commercial", "MarineRevenueSAR", 1e6),
    "Storage Revenue (SAR m)": lambda F: ssum(F, "Commercial", "StorageRevenueSAR", 1e6),
    "Discounts (SAR m)": lambda F: ssum(F, "Commercial", "DiscountSAR", 1e6),
    "Customer Tonnes (k)": lambda F: ssum(F, "Commercial", "Tonnes", 1e3),
    "Active Customers": lambda F: col(F, "Commercial", "CustomerID").nunique(),
    "Capex Actual (SAR m)": lambda F: ssum(F, "Capex", "ActualSAR", 1e6),
    "Capex Budget (SAR m)": lambda F: ssum(F, "Capex", "BudgetSAR", 1e6),
    "Forecast TEU": lambda F: fc(F, "ContainerTEU"),
    "Forecast Revenue (SAR m)": lambda F: fc(F, "RevenueSARm"),
    "Forecast Opex (SAR m)": lambda F: fc(F, "OpexSARm"),
    "Forecast EBITDA (SAR m)": lambda F: fc(F, "EBITDASARm"),
    "Forecast Capex (SAR m)": lambda F: fc(F, "CapexSARm"),
    "Quay Length (m)": lambda F: prof(F, "P04"),
    "Depth (m)": lambda F: prof(F, "P06"),
    "Terminal 1 Status": lambda F: prof(F, "P03", "Value"),
    "Capacity (TEU)": lambda F: ssum(F, "CapacityPlan", "CapacityTEU"),
    "T1 Phase 1 Capacity (TEU)": lambda F: ssum(F, "CapacityPlan", "CapacityTEU", CapacityPlan__CapacityID="C03"),
    "T1 Target Capacity (TEU)": lambda F: ssum(F, "CapacityPlan", "CapacityTEU", CapacityPlan__CapacityID="C04"),
    "Equipment Units": lambda F: ssum(F, "Equipment", "Units"),
    "STS Cranes": lambda F: ssum(F, "Equipment", "Units", Equipment__TypeCode="STS"),
    "Contract Value (SAR m)": lambda F: ssum(F, "Contracts", "ValueSARm"),
    "Contract Value (USD m)": lambda F: ssum(F, "Contracts", "ValueUSDm"),
    "Stated Investment (SAR bn)": lambda F: None if F["Investment"].empty else col(F, "Investment", "AmountSARbn").max(),
    "Stated Investment (USD bn)": lambda F: None if F["Investment"].empty else col(F, "Investment", "AmountUSDbn").max(),
    "Avg Calls per Month": lambda F: ssum(F, "VesselCalls", "AvgCallsPerMonth"),
    "RoRo Calls per Month (Pre-war)": lambda F: ssum(F, "VesselCalls", "AvgCallsPerMonth", VesselCalls__CallID="V01"),
    "RoRo Calls (Sep 2026)": lambda F: ssum(F, "VesselCalls", "AvgCallsPerMonth", VesselCalls__CallID="V02"),
}
m = lambda n: MEASURES[n]
MEASURES.update({
    "Variance (SAR m)": lambda F: sub(m("Actual (SAR m)")(F), m("Budget (SAR m)")(F)),
    "EBITDA (SAR m)": lambda F: sub(m("Revenue (SAR m)")(F), m("Opex (SAR m)")(F)),
    "Budget EBITDA (SAR m)": lambda F: sub(m("Budget Revenue (SAR m)")(F), m("Budget Opex (SAR m)")(F)),
    "Avg Turnaround (hrs)": lambda F: div(m("Berth Hours")(F), m("Vessel Calls")(F)),
    "Handling Revenue (SAR m)": lambda F: sub(ssum(F, "Commercial", "GrossHandlingSAR", 1e6), m("Discounts (SAR m)")(F)),
    "Revenue per Tonne (SAR)": lambda F: div(ssum(F, "Commercial", "NetRevenueSAR"), ssum(F, "Commercial", "Tonnes")),
    "Capex Variance (SAR m)": lambda F: sub(m("Capex Actual (SAR m)")(F), m("Capex Budget (SAR m)")(F)),
    "Revenue vs Budget %": lambda F: div(sub(m("Revenue (SAR m)")(F), m("Budget Revenue (SAR m)")(F)), m("Budget Revenue (SAR m)")(F)),
    "Phase 1 % of Saudi Volume": lambda F: div(m("T1 Phase 1 Capacity (TEU)")(F), prof(F, "P13")),
    "Base 2031 TEU": lambda F: base31(F, m("Forecast TEU")),
    "Base 2031 Revenue (SAR m)": lambda F: base31(F, m("Forecast Revenue (SAR m)")),
    "Base 2031 EBITDA (SAR m)": lambda F: base31(F, m("Forecast EBITDA (SAR m)")),
})
MEASURES.update({
    "EBITDA Margin": lambda F: div(m("EBITDA (SAR m)")(F), m("Revenue (SAR m)")(F)),
    "EBIT (SAR m)": lambda F: sub(m("EBITDA (SAR m)")(F), m("Depreciation (SAR m)")(F)),
    "RoRo Change vs Pre-war": lambda F: div(m("RoRo Calls (Sep 2026)")(F), m("RoRo Calls per Month (Pre-war)")(F)) - 1,
    "Base 2031 EBITDA Margin": lambda F: div(m("Base 2031 EBITDA (SAR m)")(F), m("Base 2031 Revenue (SAR m)")(F)),
})
FORMATS = {mm: f for _, (_, _, ms) in b.TABLES.items() for mm, _, f in ms}
TOTAL = {}


def fmt(name, v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, str):
        return v
    f = (FORMATS.get(name) or "#,##0").split(";")[0]
    if "%" in f:
        return f"{v * 100:.{len(f.split('.')[1]) - 1 if '.' in f else 0}f}%"
    dec = len(f.split(".")[1]) if "." in f else 0
    v = 0.0 if abs(v) < 0.5 * 10 ** -dec else v
    return f"{v:,.{dec}f}" if "," in f else f"{v:.{dec}f}"


def fval(v):
    if hasattr(v, "strftime"):
        return v.strftime("%d %b %Y")
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return "" if v is None else str(v)


# ---------------------------------------------------------------- drawing
def field(p):
    f = p["field"]
    kind = "Column" if "Column" in f else "Measure"
    return kind, f[kind]["Expression"]["SourceRef"]["Entity"], f[kind]["Property"], p.get("displayName")


def title_of(v):
    try:
        return v["visualContainerObjects"]["title"][0]["properties"]["text"]["expr"]["Literal"]["Value"][1:-1].replace("''", "'")
    except KeyError:
        return None


def frame_with(F, keys):
    for n, d in F.items():
        if all(k in d.columns for k in keys):
            return d
    raise KeyError(keys)


def panel(fig, x, y, w, h, title=None):
    ax = fig.add_axes([x / W, 1 - (y + h) / H, w / W, h / H])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    ax.add_patch(FancyBboxPatch((0.004, 0.01), 0.992, 0.98, boxstyle="round,pad=0,rounding_size=0.012",
                                fc="white", ec="#DDE2E8", lw=0.8, transform=ax.transAxes))
    if title:
        ax.text(0.02, 1 - 14 / h, title, fontsize=7.5, color=NAVY, weight="bold", va="center")
    return ax


def draw_chart(fig, F, v, x, y, w, h):
    panel(fig, x, y, w, h, title_of(v))
    ax = fig.add_axes([(x + 44) / W, 1 - (y + h - 26) / H, (w - 60) / W, (h - 62) / H])
    qs = v["query"]["queryState"]
    _, ct, cc, _ = field(qs["Category"]["projections"][0])
    ys = [field(p)[2] for p in qs["Y"]["projections"]]
    ser = field(qs["Series"]["projections"][0]) if "Series" in qs else None
    key = f"{ct}.{cc}"
    cats = sorted(frame_with(F, [key])[key].dropna().unique())
    if ser:
        skey = f"{ser[1]}.{ser[2]}"
        svals = sorted(frame_with(F, [skey])[skey].dropna().unique())
        series = {str(s): [MEASURES[ys[0]](flt(flt(F, ct, cc, c), ser[1], ser[2], s)) for c in cats] for s in svals}
        series = {k: vals for k, vals in series.items() if any(x is not None for x in vals)}
    else:
        series = {n: [MEASURES[n](flt(F, ct, cc, c)) for c in cats] for n in ys}
    vt = v["visualType"]
    sort = v["query"].get("sortDefinition")
    if sort and "Measure" in sort["sort"][0]["field"] and not ser:
        first = list(series.values())[0]
        order = sorted(range(len(cats)), key=lambda i: first[i] or 0, reverse=sort["sort"][0]["direction"] == "Descending")
        cats = [cats[i] for i in order]
        series = {k: [vals[i] for i in order] for k, vals in series.items()}
    is_date = hasattr(cats[0], "year") and not isinstance(cats[0], (int, float))
    xs = list(cats) if is_date else list(range(len(cats)))
    num = lambda vals: [0 if x is None else x for x in vals]
    horizontal = vt in ("barChart", "clusteredBarChart")
    n = len(series)
    if vt == "lineChart":
        for i, (k, vals) in enumerate(series.items()):
            ax.plot(xs, vals, color=PALETTE[i % 10], lw=1.4, label=k, marker="o" if len(xs) < 8 else None, ms=3)
    elif horizontal:
        hb = 0.8 / n if vt == "clusteredBarChart" else 0.7
        for i, (k, vals) in enumerate(series.items()):
            pos = [p + (i - (n - 1) / 2) * hb for p in xs] if vt == "clusteredBarChart" else xs
            bars = ax.barh(pos, num(vals), height=hb, color=PALETTE[i % 10], label=k)
            ax.bar_label(bars, labels=[fmt(ys[min(i, len(ys) - 1)], x) for x in vals], fontsize=6, padding=2, color=GREY)
        ax.invert_yaxis()
        ax.set_yticks(xs)
        ax.set_yticklabels([textwrap.shorten(str(c), 40, placeholder="...") for c in cats], fontsize=6.5)
        ax.set_position([(x + 230) / W, 1 - (y + h - 22) / H, (w - 280) / W, (h - 58) / H])
    else:
        width = 22 if is_date else 0.7
        clustered = vt == "clusteredColumnChart"
        bottom = [0.0] * len(xs)
        for i, (k, vals) in enumerate(series.items()):
            if clustered:
                bw = 0.8 / n
                bars = ax.bar([p + (i - (n - 1) / 2) * bw for p in xs], num(vals), width=bw, color=PALETTE[i % 10], label=k)
                ax.bar_label(bars, labels=[fmt(ys[min(i, len(ys) - 1)], x) for x in vals], fontsize=6, padding=1, color=GREY)
            else:
                bars = ax.bar(xs, num(vals), width=width, bottom=bottom, color=PALETTE[i % 10], label=k)
                bottom = [a + c for a, c in zip(bottom, num(vals))]
                if n == 1 and len(xs) < 10:
                    ax.bar_label(bars, labels=[fmt(ys[0], x) for x in vals], fontsize=6, padding=1, color=GREY)
    if is_date:
        ax.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[1, 7]))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %y"))
    elif not horizontal:
        ax.set_xticks(xs)
        ax.set_xticklabels([textwrap.fill(fval(c), 16) for c in cats], fontsize=6.5)
    if FORMATS.get(ys[0], "").endswith("%"):
        (ax.xaxis if horizontal else ax.yaxis).set_major_formatter(matplotlib.ticker.PercentFormatter(1.0, decimals=0))
    else:
        (ax.xaxis if horizontal else ax.yaxis).set_major_formatter(matplotlib.ticker.StrMethodFormatter("{x:,.0f}"))
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    for sp in ("left", "bottom"):
        ax.spines[sp].set_color("#C9CED4")
    ax.tick_params(colors=GREY, length=2, labelsize=6.5)
    ax.grid(axis="x" if horizontal else "y", color="#EEF0F3", lw=0.6)
    ax.set_axisbelow(True)
    if n > 1:
        ax.legend(fontsize=6, frameon=False, ncol=min(n, 6 if max(len(k) for k in series) < 16 else 3), loc="lower left", bbox_to_anchor=(0, 1.0), handlelength=1.2,
                  columnspacing=1.0)


def draw_table(fig, F, v, x, y, w, h):
    ax = panel(fig, x, y, w, h, title_of(v))
    fields = [field(p) for p in v["query"]["queryState"]["Values"]["projections"]]
    cols = [(t, c) for k, t, c, _ in fields if k == "Column"]
    keys = [f"{t}.{c}" for t, c in cols]
    base = frame_with(F, keys)
    extra = [f"{t}.{b.SORT_BY[(t, c)]}" for t, c in cols if (t, c) in b.SORT_BY]
    groups = base[keys + [e for e in extra if e not in keys]].drop_duplicates()
    rows = []
    for _, g in groups.iterrows():
        Fg = F
        for (t, c), k in zip(cols, keys):
            Fg = flt(Fg, t, c, g[k])
        vals = [g[f"{t}.{c}"] if k == "Column" else MEASURES[c](Fg) for k, t, c, _ in fields]
        if all(k == "Column" for k, *_ in fields) or any(vals[i] is not None for i, f in enumerate(fields) if f[0] == "Measure"):
            rows.append((vals, g))
    sort = v["query"].get("sortDefinition")
    if sort:
        sf = sort["sort"][0]["field"]
        kind = "Column" if "Column" in sf else "Measure"
        name = sf[kind]["Property"]
        idx = [f[2] for f in fields].index(name)
        skey = (lambda r: r[1][extra[0]]) if kind == "Column" and extra else (lambda r: (r[0][idx] is None, r[0][idx] or 0))
        rows.sort(key=skey, reverse=sort["sort"][0]["direction"] == "Descending")
    # column widths by content
    texts = [[fmt(f[2], val) if f[0] == "Measure" else fval(val) for f, val in zip(fields, vals)] for vals, _ in rows]
    heads = [f[3] or f[2] for f in fields]
    lens = [min(70, max([len(hd)] + [len(r[i]) for r in texts])) + 2 for i, hd in enumerate(heads)]
    tot = sum(lens)
    px = [(w - 24) * l / tot for l in lens]
    y0, rh = 34, 15.5
    xs = [12 + sum(px[:i]) for i in range(len(px))]

    def cell(i, yy, s, bold=False, num=False):
        chars = int(px[i] / 4.1)
        s = s if len(s) <= chars else s[:max(1, chars - 1)] + "..."
        ax.text((xs[i] + (px[i] - 8 if num else 0)) / w, 1 - yy / h, s, fontsize=6.3, va="center",
                ha="right" if num else "left", weight="bold" if bold else "normal", color=NAVY if bold else "#222222")

    for i, hd in enumerate(heads):
        cell(i, y0, hd, True, fields[i][0] == "Measure")
    ax.plot([12 / w, (w - 12) / w], [1 - (y0 + 8) / h] * 2, color=NAVY, lw=0.8)
    shown = int((h - y0 - 14) / rh)
    for r, row in enumerate(texts[:shown]):
        yy = y0 + 8 + rh * (r + 0.6)
        if r % 2:
            ax.add_patch(plt.Rectangle((10 / w, 1 - (yy + rh / 2) / h), (w - 20) / w, rh / h, fc="#F5F7F9", ec="none"))
        for i, s in enumerate(row):
            cell(i, yy, s, num=fields[i][0] == "Measure")
    if len(texts) > shown:
        ax.text(0.985, 6 / h, f"{len(texts) - shown} more rows (scroll in Power BI)", fontsize=5.5, ha="right", color=GREY, style="italic")


def draw(fig, F, v, x, y, w, h):
    vt = v["visualType"]
    if vt == "textbox":
        run = v["objects"]["general"][0]["properties"]["paragraphs"][0]["textRuns"][0]
        st = run["textStyle"]
        fig.text(x / W, 1 - (y + h / 2) / H, run["value"], fontsize=float(st["fontSize"][:-2]) * 0.72, color=st["color"],
                 weight=st.get("fontWeight", "normal"), va="center")
    elif vt == "card":
        _, _, name, label = field(v["query"]["queryState"]["Values"]["projections"][0])
        ax = panel(fig, x, y, w, h)
        val = fmt(name, MEASURES[name](F))
        if h < 70:
            ax.text(0.015, 0.5, label + ":", fontsize=7.5, color=GREY, va="center")
            ax.text(0.12, 0.5, val, fontsize=10, color="#C00000", weight="bold", va="center")
        else:
            ax.text(0.5, 0.60, val, fontsize=15 if len(val) < 9 else 12, color=NAVY, weight="bold", ha="center", va="center")
            ax.text(0.5, 0.20, textwrap.fill(label, int(w / 5.2)), fontsize=6.5, color=GREY, ha="center", va="center")
    elif vt == "slicer":
        ax = panel(fig, x, y, w, h)
        ax.text(0.06, 0.74, title_of(v), fontsize=6.5, color=GREY, va="center")
        ax.add_patch(plt.Rectangle((0.06, 0.14), 0.88, 0.40, fc="white", ec="#B8BFC7", lw=0.7))
        ax.text(0.10, 0.34, "All", fontsize=7, va="center")
        ax.text(0.90, 0.34, "v", fontsize=6, va="center", ha="right", color=GREY)
    elif vt == "tableEx":
        draw_table(fig, F, v, x, y, w, h)
    else:
        draw_chart(fig, F, v, x, y, w, h)


def main(xlsx, out):
    F = load(xlsx)
    total = MEASURES["Net Revenue (SAR m)"](F)
    MEASURES["Revenue Share"] = lambda Fg: div(MEASURES["Net Revenue (SAR m)"](Fg), total)
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    for i, (name, visuals) in enumerate(b.PAGES, 1):
        fig = plt.figure(figsize=(W / 100, H / 100), dpi=200)
        fig.patch.set_facecolor("#F3F5F8")
        for (x, y, w, h), v in visuals:
            draw(fig, F, v, x, y, w, h)
        fig.text(0.995, 0.006, "Preview rendered from the report layout and workbook data - not a Power BI screenshot",
                 fontsize=5.5, color="#8A929B", ha="right")
        fn = out / f"{i:02d}_{re.sub(r'[^A-Za-z0-9]+', '_', name).strip('_')}.png"
        fig.savefig(fn, facecolor=fig.get_facecolor())
        plt.close(fig)
        print(fn)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
