import openpyxl
from openpyxl.chart import LineChart, Reference

from failbook import kit
from failbook import xlsx as X
from failbook.checks import Inconclusive, require, rule

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
    "The Revenue and Costs series cover January to December, aligned with the month labels",
    kind="structure",
    title_zh="Revenue 和 Costs 两条系列覆盖 1 月到 12 月，并与月份标签对齐",
    requirement="The existing chart must show all twelve months, January to December, for both Revenue and Costs.",
)
def chart_covers_all_months(ctx):
    for chart in _titled_charts(ctx):
        _check_chart(ctx, chart)


SERIES = ((2, "Revenue"), (3, "Costs"))


def _check_chart(ctx, chart):
    wb, vals = ctx.wb(), ctx.values().ws(SHEET)
    found, unresolved = {}, []
    for s in chart["series"]:
        rng = X.resolve_range(s["val"], wb, SHEET)
        if rng is None:
            unresolved.append(_describe(wb, s["val"]))
        elif rng[0] == SHEET and rng[2] == rng[4]:
            found.setdefault(rng[2], (s, rng))
    for col, label in SERIES:
        if col not in found:
            if unresolved:
                raise Inconclusive(
                    f"no chart series resolves to a fixed range in the {label} column. The checker resolves "
                    "plain ranges and names that stand for a fixed range, not "
                    f"{'; '.join(unresolved)}, so whether the chart shows January to December is not checked")
            require(False, f"the chart has no {label} series (series: {[s['val'] for s in chart['series']]})")
        s, (_, r1, _, r2, _) = found[col]
        where = s["val"]
        first = X.show(vals.cell(r1, col).value)
        require(r1 == 2, f"the {label} series ({where}) starts at row {r1} ({first}"
                + (", the header" if r1 == 1 else "") + "), not at row 2 (Jan)")
        require(r2 >= 13, f"the {label} series covers rows {r1}-{r2}, not 2-13 (Jan-Dec)")
        extra = [r for r in range(14, r2 + 1) if vals.cell(r, col).value not in (None, "")]
        require(not extra, f"the {label} series ({where}) also plots row {extra[0] if extra else ''}, which is not a month")
        _check_labels(wb, vals, label, s)
        _check_name(ctx, wb, label, s)


def _describe(wb, ref):
    """What an unresolved series reference is, for an inconclusive message."""
    if not ref:
        return "a series whose values are typed into the chart"
    name = ref.split("!", 1)[1] if ref.startswith("[0]!") else ref
    for key, defn in wb.defined_names.items():
        if key.casefold() == name.casefold():
            return f"{ref} (the name {key} = {defn.attr_text})"
    return ref


def _check_labels(wb, vals, label, s):
    if not s["cat"]:
        require(False, f"the {label} series has no month labels")
    cat = X.resolve_range(s["cat"], wb, SHEET)
    if cat is None:
        raise Inconclusive(f"the month labels of the {label} series ({s['cat']}) do not resolve to a fixed range "
                           "the checker can read; they are not checked")
    sheet, c1_row, c1, c2_row, c2 = cat
    require(sheet == SHEET, f"the month labels of the {label} series come from sheet {sheet!r}, not {SHEET}")
    if c1 != c2:
        raise Inconclusive(f"the month labels of the {label} series span several columns ({s['cat']}); "
                           "multi-level labels are not checked")
    require(c1_row == 2, f"the month labels of the {label} series ({s['cat']}) start at row {c1_row}, not row 2, "
            "so they do not line up with the values")
    require(c2_row >= 13, f"the month labels of the {label} series cover rows {c1_row}-{c2_row}, not 2-13 (Jan-Dec)")
    for r, month in enumerate(MONTHS, start=2):
        got = vals.cell(r, c1).value
        require(X.norm_text(got) == X.norm_text(month),
                f"the month labels of the {label} series ({s['cat']}) read {X.show(got)} in row {r}, not {month!r}")


def _check_name(ctx, wb, label, s):
    if s["name"]:
        ref = X.resolve_range(s["name"], wb, SHEET)
        if ref is None or ref[1] != ref[3] or ref[2] != ref[4]:
            raise Inconclusive(f"the name of the {label} series comes from {s['name']}, which the checker cannot "
                               "resolve to one cell; the legend text is not checked")
        shown, source = ctx.values().ws(ref[0]).cell(ref[1], ref[2]).value, f"from {s['name']}"
    elif s.get("name_text"):
        shown, source = s["name_text"], "typed into the chart"
    else:
        require(False, f"the {label} series has no name, so the legend no longer says which line is {label} "
                "(the original chart took the names from the header row)")
    require(X.norm_text(shown) == X.norm_text(label),
            f"the series that plots the {label} column is named {X.show(shown)} ({source}), not {label!r}")
