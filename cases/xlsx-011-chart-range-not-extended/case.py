import posixpath
import re
import zipfile
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.chart import LineChart, Reference

from failbook import kit
from failbook import xlsx as X
from failbook.checks import Fail, Inconclusive, require, rule

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
    ctx.wb()  # a file that is not a workbook fails here, with the same message as the other rules
    try:
        charts = [c for c in X.charts(ctx.path()) if c["sheet"] == SHEET]
        typed_labels = _typed_labels(ctx.path())
    except (zipfile.BadZipFile, ET.ParseError, KeyError) as exc:
        # A broken package or chart part is the agent's file, not the checker; any other KeyError is a bug.
        if isinstance(exc, KeyError) and "There is no item named" not in str(exc):
            raise
        raise Fail(f"the charts in {ctx.primary!r} cannot be read: {exc}") from exc
    for chart, typed in zip(charts, typed_labels, strict=True):
        for s, labels in zip(chart["series"], typed, strict=True):
            s["cat_typed"] = labels
    return charts


def _titled_charts(ctx):
    charts = _charts(ctx)
    require(charts, f"there is no chart on the {SHEET} sheet any more")
    title = X.norm_text(ctx.meta["title"])
    matching = [c for c in charts if X.norm_text(c["title"]) == title]
    require(matching, f"no chart titled {ctx.meta['title']!r}; found titles {_clip([c['title'] for c in charts])}")
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
    require(chart["kinds"], "the chart's plot area has no chart group, so it draws no lines and is no longer a line chart")
    require("line" in chart["kinds"], f"the chart is now a {'/'.join(chart['kinds'])} chart, not a line chart")


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
    """Every part of the chart the checker can tie to a line or to the axis is checked before anything is called
    inconclusive: a part it cannot read never hides a part it can show to be wrong."""
    wb, vals = ctx.wb(), ctx.values().ws(SHEET)
    found, elsewhere, unresolved = {}, [], []
    for s in chart["series"]:
        rng = X.resolve_range(s["val"], wb, SHEET)
        if rng is None:
            unresolved.append(s)
        elif _on_sheet(wb, rng[0]) and rng[2] == rng[4] and rng[2] in dict(SERIES):
            found.setdefault(rng[2], (s, rng))
        else:
            elsewhere.append((s, rng))
    missing = [label for col, label in SERIES if col not in found]
    # A series over another range (a helper sheet or column) is the Revenue or Costs line when the cells it plots
    # hold exactly that line's January to December figures.
    by_figures = {}
    for s, rng in elsewhere if missing else []:
        plotted = _plotted(ctx, rng)
        label = next((label for label in missing if label not in by_figures and plotted is not None
                      and len(plotted) == 12 and all(map(X.close, plotted, _figures(ctx, label)))), None)
        if label:
            by_figures[label] = s
    missing = [label for label in missing if label not in by_figures]
    if missing and not unresolved:
        require(False, f"no series of the chart is over the {missing[0]} column of {SHEET} or plots the January to "
                f"December {missing[0]} figures (series: {_clip([s['val'] for s in chart['series']])})")
    described = "; ".join(_describe(wb, s["val"]) for s in unresolved)
    require(len(missing) <= len(unresolved),
            f"no series the checker can resolve is over the {' or '.join(missing)} column of {SHEET} or plots the "
            f"January to December {' or '.join(missing)} figures, and only one series cannot be resolved "
            f"({described}); one series draws one line, so the chart cannot show both {' and '.join(missing)} for "
            "January to December")
    for label in missing:
        _check_unresolved_names(ctx, wb, label, unresolved)
    open_points = []
    if missing:
        open_points.append(
            f"no chart series resolves to a fixed range in the {' or '.join(missing)} column. The checker resolves "
            f"plain ranges and names that stand for a fixed range, not {described}, so whether "
            f"{'that line shows' if len(missing) == 1 else 'those lines show'} January to December is not checked")
    for col, label in SERIES:
        if col in found:
            s, rng = found[col]
            _check_values(vals, label, col, s, rng)
            _note_inconclusive(open_points, _check_name, ctx, wb, label, s)
        elif label in by_figures:
            s = by_figures[label]
            _note_inconclusive(open_points, _check_name, ctx, wb, label, s, f"the {label} figures ({s['val']})")
    _note_inconclusive(open_points, _check_labels, ctx, wb, vals, chart)
    if open_points:
        raise Inconclusive(". ".join(p[0].upper() + p[1:] for p in open_points))


def _figures(ctx, label):
    return [m[label.lower()] for m in ctx.meta["existing"] + ctx.meta["new"]]


def _on_sheet(wb, sheet):
    """Whether a reference's sheet is Monthly. Sheet names in references match in any letter case, as in Excel and
    LibreOffice ('monthly'!$B$2:$B$13 is the Revenue column)."""
    return X.sheet_named(wb, sheet) == SHEET


def _plotted(ctx, rng):
    """The recalculated values of the cells a chart range covers, in order and without trailing empty cells; None
    when the range is neither one row nor one column, or its sheet does not exist."""
    sheet, r1, c1, r2, c2 = rng
    wb = ctx.values().wb
    name = X.sheet_named(wb, sheet)
    ws = wb[name] if name else None
    if ws is None or (r1 != r2 and c1 != c2):
        return None
    cells = ([(r, c1) for r in range(r1, min(r2, ws.max_row) + 1)] if c1 == c2
             else [(r1, c) for c in range(c1, min(c2, ws.max_column) + 1)])
    got = [getattr(ws._cells.get(rc), "value", None) for rc in cells]
    while got and got[-1] in (None, ""):
        got.pop()
    return got


def _check_unresolved_names(ctx, wb, label, unresolved):
    """When no series the checker can read draws the ``label`` line, only a series it cannot resolve can, and the
    legend must still name that line. Fails only when every such series has a name the checker can read (or a name
    that points at a sheet the workbook does not have) and none is ``label``."""
    shown = []
    for s in unresolved:
        try:
            got = _series_name(ctx, wb, s, "a series")
        except Inconclusive:
            return
        except _NameSheetMissing as exc:
            shown.append(f"{_describe(wb, s['val'])} takes its name from {exc.where}")
            continue
        if got is not None and X.norm_text(got[0]) == X.norm_text(label):
            return
        shown.append(f"{_describe(wb, s['val'])} " + ("has no name" if got is None
                     else f"is named {_clip(X.show(got[0]))} ({got[1]})"))
    require(False, f"no series the checker can resolve is over the {label} column of {SHEET} or plots the January to "
            f"December {label} figures, so only a series it cannot resolve can draw the {label} line, and "
            f"{'that series is not' if len(unresolved) == 1 else 'none of those series is'} named {label!r} "
            f"({'; '.join(shown)}); the legend does not label the {label} line")


def _note_inconclusive(open_points, check, *args):
    try:
        check(*args)
    except Inconclusive as exc:
        open_points.append(str(exc))


def _check_values(vals, label, col, s, rng):
    _, r1, _, r2, _ = rng
    where = s["val"]
    first = _clip(X.show(vals.cell(r1, col).value))
    require(r1 == 2, f"the {label} series ({where}) starts at row {r1} ({first}"
            + (", the header" if r1 == 1 else "") + "), not at row 2 (Jan)")
    require(r2 >= 13, f"the {label} series covers rows {r1}-{r2}, not 2-13 (Jan-Dec)")
    extra = [r for r in range(14, min(r2, vals.max_row) + 1) if vals.cell(r, col).value not in (None, "")]
    require(not extra, f"the {label} series ({where}) also plots row {extra[0] if extra else ''}, which is not a month")


def _describe(wb, ref):
    """What an unresolved chart reference is, for an inconclusive message."""
    if not ref:
        return "a series whose values are typed into the chart"
    name = ref.split("!", 1)[1] if ref.startswith("[0]!") else ref
    for key, defn in wb.defined_names.items():
        if key.casefold() == name.casefold():
            return f"{_clip(ref)} (the name {_clip(key)} = {_clip(defn.attr_text)})"
    return _clip(ref)


def _check_labels(ctx, wb, vals, chart):
    """The month labels on the category axis. A line chart has one category axis, drawn from the first series
    that has category labels (a range or labels typed into the chart); labels on later series are not drawn."""
    source = next((s for s in chart["series"] if s["cat"] or s.get("cat_typed") is not None), None)
    require(source, "no series of the chart has category labels, so the axis does not show the month names")
    typed = source.get("cat_typed")
    if isinstance(typed, int):
        require(False, f"the month labels typed into the chart have {typed if typed < _HUGE else 'at least 10^18'} "
                "points, not 12")
    if typed is not None:
        require([X.norm_text(v) for v in typed] == [X.norm_text(m) for m in MONTHS],
                f"the month labels typed into the chart read {_clip(typed)}, not Jan to Dec as in the Month column")
        return
    ref = source["cat"]
    cat = X.resolve_range(ref, wb, SHEET)
    if cat is None:
        raise Inconclusive(f"the month labels come from {_describe(wb, ref)}, which does not resolve to a fixed "
                           "range the checker can read; they are not checked")
    sheet, c1_row, c1, c2_row, c2 = cat
    if c1 != c2 and c1_row != c2_row:
        raise Inconclusive(f"the month labels ({ref}) span several rows and columns; multi-level labels are not "
                           "checked")
    if not _on_sheet(wb, sheet) or c1 != c2:
        # Labels from a helper sheet, or along a row: what the axis shows is the text of those cells.
        got = _plotted(ctx, cat)
        require(got is not None, f"the month labels ({ref}) come from sheet {sheet!r}, which does not exist")
        require([X.norm_text(v) for v in got] == [X.norm_text(m) for m in MONTHS],
                f"the month labels ({ref}) read {_clip(got)}, not Jan to Dec")
        return
    require(c1_row == 2, f"the month labels ({ref}) start at row {c1_row}, not row 2, "
            "so they do not line up with the values")
    require(c2_row >= 13, f"the month labels ({ref}) cover rows {c1_row}-{c2_row}, not 2-13 (Jan-Dec)")
    for r, month in enumerate(MONTHS, start=2):
        got = vals.cell(r, c1).value
        require(X.norm_text(got) == X.norm_text(month),
                f"the month labels ({ref}) read {_clip(X.show(got))} in row {r}, not {month!r}")


class _NameSheetMissing(Fail):
    """The name of a series points at a sheet the workbook does not have; ``where`` is that reference and sheet."""

    def __init__(self, message, where):
        super().__init__(message)
        self.where = where


def _series_name(ctx, wb, s, what):
    """The legend text of a series and where it comes from, as ``(text, source)``; None when the series has no
    name. Inconclusive when the name points somewhere the checker cannot read; _NameSheetMissing (a Fail) when it
    points at a sheet the workbook does not have. Sheet names match in any letter case."""
    if s["name"]:
        ref = X.resolve_range(s["name"], wb, SHEET)
        if ref is None or ref[1] != ref[3] or ref[2] != ref[4]:
            raise Inconclusive(f"the name of {what} comes from {_clip(s['name'])}, which the checker cannot "
                               "resolve to one cell; the legend text is not checked")
        vals = ctx.values().wb
        sheet = X.sheet_named(vals, ref[0])
        if sheet is None:
            where = f"{_clip(s['name'])}, but the workbook has no sheet {_clip(ref[0])!r}"
            raise _NameSheetMissing(f"the name of {what} comes from {where} (sheets: "
                                    f"{_clip(', '.join(vals.sheetnames))}), so the legend cannot read the line's "
                                    "name from it", where)
        return vals[sheet].cell(ref[1], ref[2]).value, f"from {_clip(s['name'])}"
    if s.get("name_text"):
        return s["name_text"], "typed into the chart"
    return None


def _check_name(ctx, wb, label, s, plots=None):
    got = _series_name(ctx, wb, s, f"the {label} series")
    require(got, f"the {label} series has no name, so the legend no longer says which line is {label} "
            "(the original chart took the names from the header row)")
    shown, source = got
    require(X.norm_text(shown) == X.norm_text(label),
            f"the series that plots {plots or f'the {label} column'} is named {_clip(X.show(shown))} ({source}), "
            f"not {label!r}")


def _clip(value, limit=200):
    """``value`` as text for a message, cut to ``limit`` characters (lists keep their first 12 items)."""
    if isinstance(value, (list, tuple)) and len(value) > 12:
        value = [*value[:12], f"... {len(value) - 12} more"]
    text = value if isinstance(value, str) else str(value)
    return text if len(text) <= limit else text[:limit] + "..."


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


_INDEX = re.compile(r"[0-9]{1,18}")
_HUGE = 10 ** 18


def _number(text):
    """A point count or index written in the chart: an int, None when it is not a plain number, and ``_HUGE``
    for a number of 19 or more significant digits (leading zeros do not count)."""
    text = (text or "").strip()
    if not (text.isascii() and text.isdigit()):
        return None
    text = text.lstrip("0") or "0"
    return int(text) if _INDEX.fullmatch(text) else _HUGE


def _series_typed_labels(root):
    """Per series, the 12 labels typed into the chart, None when there are none, or the number of points they
    declare (``c:ptCount`` or the highest ``c:pt`` index) when that is not 12. Only the first 12 points are ever
    read, so a declared count of millions costs nothing."""
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
            count = lit.find("c:ptCount", _NS)
            size = (_number(count.get("val")) if count is not None else None) or 0
            points = {}
            for pt in lit.findall("c:pt", _NS):
                idx = _number(pt.get("idx"))
                if idx is None:
                    continue
                size = max(size, idx + 1)
                if idx < 12:
                    points[idx] = pt.findtext("c:v", "", _NS)
            out.append([points.get(i, "") for i in range(12)] if size == 12 else size)
    return out
