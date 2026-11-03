import openpyxl
from openpyxl.chart import LineChart, Reference

from failbook import kit
from failbook import xlsx as X
from failbook.checks import require, rule

SHEET = "Monthly"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def params(seed):
    r = kit.rng(seed, "xlsx-011")
    base = r.randrange(80_000, 160_000, 1000)
    rows = []
    for m in MONTHS:
        revenue = int(base * r.uniform(0.85, 1.3))
        rows.append({"month": m, "revenue": revenue, "costs": int(revenue * r.uniform(0.55, 0.85))})
    title = r.choice(["Revenue vs Costs 2024", "Monthly Revenue and Costs", "2024 Revenue vs Costs"])
    new = rows[10:]
    table = kit.md_table(["Month", "Revenue", "Costs"], [(m["month"], m["revenue"], m["costs"]) for m in new])
    return {"existing": rows[:10], "new": new, "title": title, "new_md": table}


def build(meta, outdir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.append(["Month", "Revenue", "Costs"])
    kit.style_header(ws, 1, 3)
    for i, m in enumerate(meta["existing"], start=2):
        ws.append([m["month"], m["revenue"], m["costs"]])
        ws.cell(i, 2).number_format = ws.cell(i, 3).number_format = kit.INT_FMT
    last = 1 + len(meta["existing"])
    chart = LineChart()
    chart.title = meta["title"]
    chart.y_axis.title = "Amount"
    chart.add_data(Reference(ws, min_col=2, max_col=3, min_row=1, max_row=last), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last))
    chart.width, chart.height = 18, 9
    ws.add_chart(chart, "E2")
    kit.set_widths(ws, {"A": 8, "B": 12, "C": 12})
    wb.save(outdir / "monthly.xlsx")


def _charts(ctx):
    return [c for c in X.charts(ctx.path()) if c["sheet"] == SHEET]


def _titled_charts(ctx):
    charts = _charts(ctx)
    require(charts, f"there is no chart on the {SHEET} sheet any more")
    title = X.norm_text(ctx.meta["title"])
    matching = [c for c in charts if X.norm_text(c["title"]) == title]
    require(matching, f"no chart titled {ctx.meta['title']!r}; found titles {[c['title'] for c in charts]}")
    return matching


@rule(
    "new-months-added",
    "November and December are added directly below October with the given figures",
    kind="content",
    title_zh="11 月和 12 月的数据紧接在 10 月下方，数值正确",
    requirement="Add the November and December figures directly below October.",
)
def new_months_added(ctx):
    vals = ctx.values().ws(SHEET)
    for i, m in enumerate(ctx.meta["new"], start=12):
        got = [vals.cell(i, c).value for c in range(1, 4)]
        require(X.norm_text(got[0]) == X.norm_text(m["month"]), f"row {i} should be {m['month']}, found {X.show(got[0])}")
        require(X.close(got[1], m["revenue"]) and X.close(got[2], m["costs"]), f"{m['month']} figures are {got[1:]}")


@rule(
    "existing-data-unchanged",
    "January to October are unchanged",
    kind="preserve",
    title_zh="1 月至 10 月的数据保持不变",
    requirement="Do not change anything the task does not ask you to change.",
)
def existing_data_unchanged(ctx):
    vals = ctx.values().ws(SHEET)
    for i, m in enumerate(ctx.meta["existing"], start=2):
        got = [vals.cell(i, c).value for c in range(1, 4)]
        require(got == [m["month"], m["revenue"], m["costs"]], f"row {i} changed: {got}")


@rule(
    "chart-present",
    "The line chart with the original title is still on the sheet",
    kind="structure",
    title_zh="原标题的折线图仍在工作表上",
    requirement="The existing chart must ... Keep it a line chart with the same title.",
)
def chart_present(ctx):
    charts = _titled_charts(ctx)
    require(len(charts) == 1, f"there are {len(charts)} charts titled {ctx.meta['title']!r}; expected the one existing chart")
    chart = charts[0]
    require("line" in chart["kinds"], f"the chart is now a {'/'.join(chart['kinds']) or 'unknown'} chart, not a line chart")


@rule(
    "chart-covers-all-months",
    "Both series cover January to December",
    kind="structure",
    title_zh="两条数据系列都覆盖 1 月到 12 月",
    requirement="The existing chart must show all twelve months, January to December, for both Revenue and Costs.",
)
def chart_covers_all_months(ctx):
    for chart in _titled_charts(ctx):
        _check_chart(chart)


def _check_chart(chart):
    found = {}
    for s in chart["series"]:
        rng = X.parse_range(s["val"] or "")
        if rng and rng[0] == SHEET and rng[2] == rng[4]:
            found[rng[2]] = (rng, X.parse_range(s["cat"] or ""))
    for col, label in ((2, "Revenue"), (3, "Costs")):
        require(col in found, f"the chart has no {label} series (series: {[s['val'] for s in chart['series']]})")
        (sheet, r1, _, r2, _), cat = found[col]
        require(r1 <= 2 and r2 >= 13, f"the {label} series covers rows {r1}-{r2}, not 2-13 (Jan-Dec)")
        require(cat is not None and cat[1] <= 2 and cat[3] >= 13, f"the {label} series' month labels do not cover Jan-Dec")
