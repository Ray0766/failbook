import openpyxl
from openpyxl.workbook.defined_name import DefinedName

from failbook import kit
from failbook import xlsx as X
from failbook.checks import require, rule

OLD, NEW, DASH = "Data", "Sales 2024", "Dashboard"
MONTHS = ["Jan", "Feb", "Mar"]
ALL_LABEL = "All regions"
AVG_LABEL = "Average per region-month"


def params(seed):
    r = kit.rng(seed, "xlsx-006")
    regions = kit.sample(r, kit.REGIONS, 4)
    rows = [{"region": reg, "month": m, "revenue": r.randrange(20000, 90000, 50)} for m in MONTHS for reg in regions]
    return {"regions": regions, "rows": rows}


def build(meta, outdir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = OLD
    ws.append(["Region", "Month", "Revenue"])
    kit.style_header(ws, 1, 3)
    for i, row in enumerate(meta["rows"], start=2):
        ws.append([row["region"], row["month"], row["revenue"]])
        ws.cell(i, 3).number_format = kit.INT_FMT
    last = 1 + len(meta["rows"])
    wb.defined_names["RevenueRange"] = DefinedName("RevenueRange", attr_text=f"{OLD}!$C$2:$C${last}")
    kit.set_widths(ws, {"A": 12, "B": 8, "C": 12})

    d = wb.create_sheet(DASH)
    d.append(["Region", "Revenue"])
    kit.style_header(d, 1, 2)
    for i, region in enumerate(meta["regions"], start=2):
        d.cell(i, 1, region)
        d.cell(i, 2, f'=SUMIFS({OLD}!C:C,{OLD}!A:A,A{i})').number_format = kit.INT_FMT
    n = len(meta["regions"])
    d.cell(n + 2, 1, ALL_LABEL).font = kit.TOTAL_FONT
    d.cell(n + 2, 2, "=SUM(RevenueRange)").number_format = kit.INT_FMT
    d.cell(n + 3, 1, AVG_LABEL)
    d.cell(n + 3, 2, f"=AVERAGE({OLD}!C2:C{last})").number_format = kit.INT_FMT
    kit.set_widths(d, {"A": 26, "B": 12})
    wb.save(outdir / "report.xlsx")


def _expected(meta, bump_row=None, bump=0):
    rows = [dict(r) for r in meta["rows"]]
    if bump_row is not None:
        rows[bump_row]["revenue"] += bump
    out = {reg: sum(r["revenue"] for r in rows if r["region"] == reg) for reg in meta["regions"]}
    out[ALL_LABEL] = sum(r["revenue"] for r in rows)
    out[AVG_LABEL] = out[ALL_LABEL] / len(rows)
    return out


def _dashboard_refs(ctx):
    ws = ctx.ws(DASH)
    refs = {}
    for label in [*ctx.meta["regions"], ALL_LABEL, AVG_LABEL]:
        hit = X.find_cell(ws, label, col=1)
        require(hit, f"Dashboard label {label!r} is missing")
        refs[label] = X.make_ref(DASH, (hit[0], 2))
    return refs


def scenarios(ctx):
    wb = ctx.wb()
    require(NEW in wb.sheetnames, f"there is no sheet named {NEW!r}")
    return {"bump": {X.make_ref(NEW, "C2"): ctx.meta["rows"][0]["revenue"] + 1000}}


@rule(
    "sheet-renamed",
    f"The sheet is now called '{NEW}' (and '{OLD}' is gone), with the same data",
    kind="content",
    title_zh=f"工作表已改名为 '{NEW}'（'{OLD}' 不再存在），数据不变",
    requirement="Rename the sheet Data to Sales 2024.",
)
def sheet_renamed(ctx):
    wb = ctx.wb()
    require(NEW in wb.sheetnames, f"no sheet named {NEW!r}; sheets are {wb.sheetnames}")
    require(OLD not in wb.sheetnames, f"a sheet named {OLD!r} still exists")
    vals = ctx.values().ws(NEW)
    for i, row in enumerate(ctx.meta["rows"], start=2):
        got = (vals.cell(i, 1).value, vals.cell(i, 2).value, vals.cell(i, 3).value)
        require(got == (row["region"], row["month"], row["revenue"]), f"row {i} of {NEW!r} changed: {got}")


@rule(
    "dashboard-values",
    "Every Dashboard figure has its original value",
    kind="value",
    title_zh="Dashboard 上每个数字都与原来相同",
    requirement="Everything on the Dashboard must keep working exactly as before: the same figures.",
)
def dashboard_values(ctx):
    values = ctx.values()
    for label, ref in _dashboard_refs(ctx).items():
        got = ctx.num(values, ref, f"Dashboard '{label}'")
        exp = _expected(ctx.meta)[label]
        require(X.close(got, exp), f"Dashboard '{label}' is {got:,.2f}, expected {exp:,.2f}")


@rule(
    "dashboard-live",
    f"Dashboard figures still update from '{NEW}'",
    kind="live",
    title_zh=f"Dashboard 数字仍随 '{NEW}' 的数据更新",
    requirement="...still calculated from the renamed sheet, so they keep updating when the sales data changes.",
)
def dashboard_live(ctx):
    values = ctx.values(scenario="bump")
    exp = _expected(ctx.meta, bump_row=0, bump=1000)
    for label, ref in _dashboard_refs(ctx).items():
        got = ctx.num(values, ref, f"Dashboard '{label}'")
        require(X.close(got, exp[label]),
                f"after adding 1,000 to {NEW}!C2, Dashboard '{label}' is {got:,.2f}, expected {exp[label]:,.2f}")
