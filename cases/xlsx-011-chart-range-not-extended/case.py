import posixpath
import zipfile
from xml.etree import ElementTree as ET

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
    charts = [c for c in X.charts(ctx.path()) if c["sheet"] == SHEET]
    for chart, typed in zip(charts, _typed_labels(ctx.path()), strict=True):
        for s, labels in zip(chart["series"], typed, strict=True):
            s["cat_typed"] = labels
    return charts


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
    """Every series and the month labels are checked before anything is called inconclusive: a part the
    checker cannot read never hides a part it can show to be wrong."""
    wb, vals = ctx.wb(), ctx.values().ws(SHEET)
    found, unresolved = {}, []
    for s in chart["series"]:
        rng = X.resolve_range(s["val"], wb, SHEET)
        if rng is None:
            unresolved.append(_describe(wb, s["val"]))
        elif rng[0] == SHEET and rng[2] == rng[4]:
            found.setdefault(rng[2], (s, rng))
    missing = [label for col, label in SERIES if col not in found]
    if missing and not unresolved:
        require(False, f"the chart has no {missing[0]} series (series: {[s['val'] for s in chart['series']]})")
    require(len(missing) <= len(unresolved),
            f"the chart has no series over the {' or '.join(missing)} column and only one series the checker "
            f"cannot resolve ({'; '.join(unresolved)}); one series draws one line, so the chart cannot show both "
            f"{' and '.join(missing)}")
    open_points = []
    if missing:
        open_points.append(
            f"no chart series resolves to a fixed range in the {' or '.join(missing)} column. The checker resolves "
            "plain ranges and names that stand for a fixed range, not "
            f"{'; '.join(unresolved)}, so whether {'that line shows' if len(missing) == 1 else 'those lines show'} "
            "January to December is not checked")
    for col, label in SERIES:
        if col in found:
            s, rng = found[col]
            _check_values(vals, label, col, s, rng)
            _note_inconclusive(open_points, _check_name, ctx, wb, label, s)
    _note_inconclusive(open_points, _check_labels, wb, vals, chart)
    if open_points:
        raise Inconclusive(". ".join(p[0].upper() + p[1:] for p in open_points))


def _note_inconclusive(open_points, check, *args):
    try:
        check(*args)
    except Inconclusive as exc:
        open_points.append(str(exc))


def _check_values(vals, label, col, s, rng):
    _, r1, _, r2, _ = rng
    where = s["val"]
    first = X.show(vals.cell(r1, col).value)
    require(r1 == 2, f"the {label} series ({where}) starts at row {r1} ({first}"
            + (", the header" if r1 == 1 else "") + "), not at row 2 (Jan)")
    require(r2 >= 13, f"the {label} series covers rows {r1}-{r2}, not 2-13 (Jan-Dec)")
    extra = [r for r in range(14, r2 + 1) if vals.cell(r, col).value not in (None, "")]
    require(not extra, f"the {label} series ({where}) also plots row {extra[0] if extra else ''}, which is not a month")


def _describe(wb, ref):
    """What an unresolved chart reference is, for an inconclusive message."""
    if not ref:
        return "a series whose values are typed into the chart"
    name = ref.split("!", 1)[1] if ref.startswith("[0]!") else ref
    for key, defn in wb.defined_names.items():
        if key.casefold() == name.casefold():
            return f"{ref} (the name {key} = {defn.attr_text})"
    return ref


def _check_labels(wb, vals, chart):
    """The month labels on the category axis. A line chart has one category axis, drawn from the first series
    that has category labels (a range or labels typed into the chart); labels on later series are not drawn."""
    source = next((s for s in chart["series"] if s["cat"] or s.get("cat_typed") is not None), None)
    require(source, "no series of the chart has category labels, so the axis does not show the month names")
    typed = source.get("cat_typed")
    if typed is not None:
        require([X.norm_text(v) for v in typed] == [X.norm_text(m) for m in MONTHS],
                f"the month labels typed into the chart read {typed}, not Jan to Dec as in the Month column")
        return
    ref = source["cat"]
    cat = X.resolve_range(ref, wb, SHEET)
    if cat is None:
        raise Inconclusive(f"the month labels come from {_describe(wb, ref)}, which does not resolve to a fixed "
                           "range the checker can read; they are not checked")
    sheet, c1_row, c1, c2_row, c2 = cat
    require(sheet == SHEET, f"the month labels ({ref}) come from sheet {sheet!r}, not {SHEET}")
    if c1 != c2:
        raise Inconclusive(f"the month labels ({ref}) span several columns; multi-level labels are not checked")
    require(c1_row == 2, f"the month labels ({ref}) start at row {c1_row}, not row 2, "
            "so they do not line up with the values")
    require(c2_row >= 13, f"the month labels ({ref}) cover rows {c1_row}-{c2_row}, not 2-13 (Jan-Dec)")
    for r, month in enumerate(MONTHS, start=2):
        got = vals.cell(r, c1).value
        require(X.norm_text(got) == X.norm_text(month),
                f"the month labels ({ref}) read {X.show(got)} in row {r}, not {month!r}")


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


_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}
_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"


def _typed_labels(path):
    """Category labels typed into the chart (``c:cat/c:strLit`` or ``c:numLit``) for each series of each chart
    on the sheet, in the order X.charts lists them; None for a series without. X.charts reports category ranges
    only, so this walks the package the same way (workbook, sheet, drawing, chart parts)."""
    out = []
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())

        def rels(part):
            base = posixpath.dirname(part)
            rels_part = posixpath.join(base, "_rels", posixpath.basename(part) + ".rels")
            if rels_part not in names:
                return {}
            found = {}
            for rel in ET.fromstring(z.read(rels_part)).iter(_REL):
                target = rel.get("Target", "")
                if rel.get("TargetMode") != "External":
                    found[rel.get("Id")] = (target.lstrip("/") if target.startswith("/")
                                           else posixpath.normpath(posixpath.join(base, target)))
            return found

        wb_rels = rels("xl/workbook.xml")
        for sheet in ET.fromstring(z.read("xl/workbook.xml")).iterfind("main:sheets/main:sheet", _NS):
            part = wb_rels.get(sheet.get(f"{{{_NS['r']}}}id"))
            if sheet.get("name") != SHEET or part not in names:
                continue
            sheet_rels = rels(part)
            for drawing in ET.fromstring(z.read(part)).iterfind("main:drawing", _NS):
                drawing_part = sheet_rels.get(drawing.get(f"{{{_NS['r']}}}id"))
                if drawing_part not in names:
                    continue
                for chart_part in rels(drawing_part).values():
                    if chart_part in names and "/charts/" in chart_part:
                        out.append(_series_typed_labels(ET.fromstring(z.read(chart_part))))
    return out


def _series_typed_labels(root):
    plot = root.find("c:chart/c:plotArea", _NS)
    out = []
    for group in plot if plot is not None else []:
        if not group.tag.split("}")[-1].endswith("Chart"):
            continue
        for ser in group.findall("c:ser", _NS):
            lit = ser.find("c:cat/c:strLit", _NS)
            if lit is None:
                lit = ser.find("c:cat/c:numLit", _NS)
            if lit is None:
                out.append(None)
                continue
            points = {int(pt.get("idx")): pt.findtext("c:v", "", _NS)
                      for pt in lit.findall("c:pt", _NS) if (pt.get("idx") or "").isdigit()}
            count = lit.find("c:ptCount", _NS)
            n = int(count.get("val")) if count is not None and (count.get("val") or "").isdigit() else 0
            out.append([points.get(i, "") for i in range(max(n, max(points, default=-1) + 1))])
    return out
