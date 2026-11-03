import datetime as dt

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from failbook import kit
from failbook import xlsx as X
from failbook.checks import require, rule

CUST, REGIONS, NOTES = "Customers", "Regions", "Notes"
HEADERS = ["ID", "Name", "City", "Tier", "Since", "Annual Spend", "Region"]
CITY_REGION = {
    "Beijing": "North", "Tianjin": "North", "Shanghai": "East", "Hangzhou": "East",
    "Shenzhen": "South", "Guangzhou": "South", "Chengdu": "West", "Wuhan": "Central",
}
TIERS = ["Gold", "Silver", "Bronze"]
TYPO, FIXED = "Bejing", "Beijing"


def params(seed):
    r = kit.rng(seed, "xlsx-012")
    n = r.randint(20, 28)
    names = kit.people(r, n)
    customers = []
    for i, name in enumerate(names):
        city = r.choice(list(CITY_REGION))
        customers.append({
            "id": f"C{2001 + i}",
            "name": name,
            "city": city,
            "tier": r.choice(TIERS),
            "since": dt.date(r.randint(2015, 2023), r.randint(1, 12), r.randint(1, 28)).isoformat(),
            "spend": round(r.uniform(1500, 90000), 2),
        })
    beijing = [c for c in customers if c["city"] == FIXED]
    while len(beijing) < 5:
        c = r.choice([c for c in customers if c["city"] != FIXED])
        c["city"] = FIXED
        beijing.append(c)
    typos = r.sample(beijing, r.randint(3, len(beijing) - 1))
    for c in typos:
        c["city"] = TYPO
    return {"customers": customers, "typo_ids": [c["id"] for c in typos]}


def build(meta, outdir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = CUST
    ws.append(HEADERS)
    kit.style_header(ws, 1, len(HEADERS))
    last = 1 + len(meta["customers"])
    for r, c in enumerate(meta["customers"], start=2):
        ws.append([c["id"], c["name"], c["city"], c["tier"], kit.from_iso(c["since"]), c["spend"],
                   f'=IFERROR(VLOOKUP(C{r},Regions!$A$2:$B${1 + len(CITY_REGION)},2,FALSE),"Unknown")'])
        ws.cell(r, 5).number_format = kit.DATE_FMT
        ws.cell(r, 6).number_format = kit.MONEY_FMT
    dv = DataValidation(type="list", formula1='"' + ",".join(TIERS) + '"', allow_blank=True)
    dv.add(f"D2:D{last}")
    ws.add_data_validation(dv)
    kit.set_widths(ws, {"A": 8, "B": 18, "C": 12, "D": 9, "E": 12, "F": 14, "G": 10})
    ws.freeze_panes = "A2"

    reg = wb.create_sheet(REGIONS)
    reg.append(["City", "Region"])
    kit.style_header(reg, 1, 2)
    for city, region in CITY_REGION.items():
        reg.append([city, region])
    kit.set_widths(reg, {"A": 12, "B": 10})

    notes = wb.create_sheet(NOTES)
    notes["A1"] = "Import log"
    notes["A1"].font = kit.TOTAL_FONT
    notes["A2"] = "2024-01-05: customers imported from the legacy CRM export."
    notes["A3"] = "Legacy CRM spelled the city as 'Bejing'; keep this note as written."
    kit.set_widths(notes, {"A": 70})
    wb.save(outdir / "customers.xlsx")


def _expected_rows(meta):
    rows = []
    for r, c in enumerate(meta["customers"], start=2):
        city = FIXED if c["city"] == TYPO else c["city"]
        rows.append((r, c, city, CITY_REGION.get(city, "Unknown")))
    return rows


@rule(
    "typo-fixed",
    "Every 'Bejing' in the City column now reads 'Beijing'",
    kind="content",
    title_zh="City 列中所有 'Bejing' 都已改为 'Beijing'",
    requirement="Correct those cells to 'Beijing'.",
)
def typo_fixed(ctx):
    vals = ctx.values().ws(CUST)
    for r, c, city, _ in _expected_rows(ctx.meta):
        got = vals.cell(r, 3).value
        require(got == city, f"City of {c['id']} (row {r}) is {X.show(got)}, expected {city!r}")


@rule(
    "customer-data-unchanged",
    "All other Customers cells are unchanged; Region is still a formula and now resolves",
    kind="preserve",
    title_zh="Customers 表其余单元格不变；Region 仍是公式，且现在能查到结果",
    requirement="Do not change anything else: keep ... the formulas ... exactly as they are.",
)
def customer_data_unchanged(ctx):
    values = ctx.values()
    vals = values.ws(CUST)
    ws = ctx.ws(CUST)
    before = ctx.input_wb(data_only=False)[CUST]
    for col in range(1, len(HEADERS) + 1):
        a, b = before.cell(1, col).value, vals.cell(1, col).value
        require(a == b, f"header {get_column_letter(col)}1 changed from {X.show(a)} to {X.show(b)}")
    last = 1 + len(ctx.meta["customers"])
    extra = [f"{get_column_letter(c)}{r}" for r, c, _ in X.used_cells(ws) if r > last or c > len(HEADERS)]
    require(not extra, f"the Customers sheet has new content outside the table: {', '.join(extra[:5])}")
    for r, c, _, region in _expected_rows(ctx.meta):
        got = [vals.cell(r, col).value for col in (1, 2, 4, 5, 6)]
        exp = [c["id"], c["name"], c["tier"], kit.from_iso(c["since"]), c["spend"]]
        require(got[:3] == exp[:3], f"row {r} changed: {got[:3]} instead of {exp[:3]}")
        require(X.as_date(got[3]) == exp[3], f"Since of {c['id']} is {X.show(got[3])}, expected {exp[3]}")
        require(X.close(got[4], exp[4]), f"Annual Spend of {c['id']} is {X.show(got[4])}, expected {exp[4]}")
        require(X.formula_text(ws.cell(r, 7).value), f"Region of {c['id']} is no longer a formula: {X.show(ws.cell(r, 7).value)}")
        got_region = ctx.value(values, X.make_ref(CUST, (r, 7)))
        require(got_region == region, f"Region of {c['id']} is {X.show(got_region)}, expected {region!r}")


@rule(
    "other-sheets-unchanged",
    "The Regions and Notes sheets are still there, unchanged",
    kind="preserve",
    title_zh="Regions 和 Notes 工作表仍在且未被修改",
    requirement="Keep the other sheets ... exactly as they are.",
)
def other_sheets_unchanged(ctx):
    wb = ctx.wb()
    original = ctx.input_wb()
    require(wb.sheetnames == original.sheetnames, f"sheets are {wb.sheetnames}, expected {original.sheetnames}")
    for name in (REGIONS, NOTES):
        before, after = original[name], wb[name]
        for row in before.iter_rows():
            for cell in row:
                got = after[cell.coordinate].value
                require(got == cell.value, f"{name}!{cell.coordinate} changed from {X.show(cell.value)} to {X.show(got)}")


@rule(
    "formatting-kept",
    "Cell formatting, column widths, the frozen header and the filter settings are kept",
    kind="structure",
    title_zh="单元格格式、列宽、冻结表头和筛选设置都保持原样",
    requirement="Keep ... the formatting ... exactly as it is.",
)
def formatting_kept(ctx):
    ws = ctx.ws(CUST)
    original = ctx.input_wb(data_only=False)
    before = original[CUST]
    wb = ctx.wb()
    for r in range(1, 2 + len(ctx.meta["customers"])):
        for col in range(1, len(HEADERS) + 1):
            a, b = before.cell(r, col), ws.cell(r, col)
            for (what, va), (_, vb) in zip(_style(a, original), _style(b, wb)):
                require(va is None or vb is None or va == vb,
                        f"{what} of {b.coordinate} changed from {_show_style(va)} to {_show_style(vb)}")
            require(a.number_format == b.number_format,
                    f"number format of {b.coordinate} changed from {a.number_format!r} to {b.number_format!r}")
    for letter in "ABCDEFG":
        wa = before.column_dimensions[letter].width or 0
        wb_ = ws.column_dimensions[letter].width or 0
        require(abs(wa - wb_) <= 1.5, f"column {letter} width changed from {wa:g} to {wb_:g}")
    require(ws.freeze_panes == before.freeze_panes, f"frozen panes changed from {before.freeze_panes} to {ws.freeze_panes}")
    require(ws.auto_filter.ref == before.auto_filter.ref,
            f"the filter changed from {before.auto_filter.ref or 'none'} to {ws.auto_filter.ref or 'none'}")
    n_before = sum(len(cf.rules) for cf in before.conditional_formatting)
    n_after = sum(len(cf.rules) for cf in ws.conditional_formatting)
    require(n_after == n_before, f"the sheet now has {n_after} conditional-format rule(s), the input had {n_before}")


def _style(cell, wb):
    """What a reader sees of a cell's style, as comparable values (None where it cannot be resolved)."""
    font, fill = cell.font, cell.fill
    return [
        ("bold", bool(font.b)),
        ("italic", bool(font.i)),
        ("underline", font.u or "none"),
        ("strikethrough", bool(font.strike)),
        ("font colour", _font_colour(font, wb)),
        ("fill", _fill(fill, wb)),
        ("horizontal alignment", cell.alignment.horizontal or "general"),
    ]


def _font_colour(font, wb):
    color = font.color
    if color is None or getattr(color, "auto", None) or (color.type == "rgb" and color.rgb == "00000000"):
        return "000000"  # automatic: the default text colour
    return X.colour_rgb(color, wb)


def _fill(fill, wb):
    if fill is None or fill.fill_type in (None, "none"):
        return "none"
    if fill.fill_type != "solid":
        return None if fill.__class__.__name__ != "PatternFill" else (fill.fill_type,)
    return X.colour_rgb(fill.fgColor, wb)


def _show_style(value):
    if value is True or value is False:
        return "on" if value else "off"
    return repr(value) if isinstance(value, str) else str(value)


@rule(
    "dropdown-kept",
    "The Tier dropdown still covers every customer row",
    kind="structure",
    title_zh="Tier 下拉列表仍覆盖所有客户行",
    requirement="Keep ... the data validation exactly as it is.",
)
def dropdown_kept(ctx):
    ws = ctx.ws(CUST)
    cells = [f"D{r}" for r in range(2, 2 + len(ctx.meta["customers"]))]
    for dv in ws.data_validations.dataValidation:
        if dv.type != "list" or not dv.formula1:
            continue
        options = {X.norm_text(o) for o in dv.formula1.strip('"').split(",")}
        if options == {X.norm_text(t) for t in TIERS} and X.ranges_cover(dv.sqref, cells):
            return
    raise AssertionError("no Tier dropdown (Gold, Silver, Bronze) covers the Tier column any more")
