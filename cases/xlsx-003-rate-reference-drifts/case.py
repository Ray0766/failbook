import openpyxl

from failbook import kit
from failbook import xlsx as X
from failbook.checks import require, rule

SHEET = "Invoice"
HEADERS = ["Item", "Qty", "Unit Price", "Net", "VAT", "Gross"]
FIRST = 4
RATES = [0.13, 0.2, 0.19, 0.08, 0.1, 0.21]


def params(seed):
    r = kit.rng(seed, "xlsx-003")
    n = r.randint(8, 12)
    lines = [
        {"item": name, "qty": r.randint(1, 25), "price": round(r.uniform(3, 250), 2)}
        for name in kit.sample(r, kit.PRODUCTS, n)
    ]
    return {"rate": r.choice(RATES), "lines": lines}


def _net(line):
    return line["qty"] * line["price"]


def build(meta, outdir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    ws["A1"] = "VAT rate"
    ws["A1"].font = kit.TOTAL_FONT
    ws["B1"] = meta["rate"]
    ws["B1"].number_format = "0%"
    for c, h in enumerate(HEADERS, start=1):
        ws.cell(3, c, h)
    kit.style_header(ws, 3, len(HEADERS))
    last = FIRST + len(meta["lines"]) - 1
    for r, line in enumerate(meta["lines"], start=FIRST):
        ws.cell(r, 1, line["item"])
        ws.cell(r, 2, line["qty"])
        ws.cell(r, 3, line["price"]).number_format = kit.MONEY_FMT
        ws.cell(r, 4, f"=B{r}*C{r}").number_format = kit.MONEY_FMT
    total = last + 1
    ws.cell(total, 1, "Total").font = kit.TOTAL_FONT
    for c, col in ((4, "D"), (5, "E"), (6, "F")):
        cell = ws.cell(total, c, f"=SUM({col}{FIRST}:{col}{last})")
        cell.font = kit.TOTAL_FONT
        cell.number_format = kit.MONEY_FMT
    kit.set_widths(ws, {"A": 18, "B": 8, "C": 11, "D": 12, "E": 11, "F": 12})
    wb.save(outdir / "invoice.xlsx")


def _rows(ctx):
    return list(range(FIRST, FIRST + len(ctx.meta["lines"])))


def _total_row(ctx):
    return FIRST + len(ctx.meta["lines"])


def _check_lines(ctx, values, rate, what=""):
    for r, line in zip(_rows(ctx), ctx.meta["lines"]):
        vat = ctx.num(values, X.make_ref(SHEET, (r, 5)), f"VAT of '{line['item']}'")
        exp = _net(line) * rate
        require(X.close(vat, exp), f"{what}VAT of '{line['item']}' (row {r}) is {vat:,.2f}, expected {exp:,.2f}")


def scenarios(ctx):
    return {"rate_up": {X.make_ref(SHEET, "B1"): round(ctx.meta["rate"] + 0.05, 4)}}


@rule(
    "vat-correct",
    "VAT = Net x rate on every line",
    kind="value",
    title_zh="每一行 VAT = Net × 税率",
    requirement="VAT = Net x the VAT rate in B1, for every line item.",
)
def vat_correct(ctx):
    _check_lines(ctx, ctx.values(), ctx.meta["rate"])


@rule(
    "gross-correct",
    "Gross = Net + VAT on every line",
    kind="value",
    title_zh="每一行 Gross = Net + VAT",
    requirement="Gross = Net + VAT, for every line item.",
)
def gross_correct(ctx):
    values = ctx.values()
    for r, line in zip(_rows(ctx), ctx.meta["lines"]):
        got = ctx.num(values, X.make_ref(SHEET, (r, 6)), f"Gross of '{line['item']}'")
        exp = _net(line) * (1 + ctx.meta["rate"])
        require(X.close(got, exp), f"Gross of '{line['item']}' (row {r}) is {got:,.2f}, expected {exp:,.2f}")


@rule(
    "vat-follows-rate",
    "Changing the rate in B1 updates the VAT of every line",
    kind="live",
    title_zh="修改 B1 的税率后，每一行的 VAT 都会更新",
    requirement="Use formulas, so that changing the rate in B1 updates every line and the totals.",
)
def vat_follows_rate(ctx):
    new_rate = round(ctx.meta["rate"] + 0.05, 4)
    _check_lines(ctx, ctx.values(scenario="rate_up"), new_rate, what=f"with B1 changed to {new_rate:.0%}, ")


@rule(
    "totals-correct",
    "The Total row shows the correct Net, VAT and Gross totals",
    kind="value",
    title_zh="Total 行的 Net、VAT、Gross 合计正确",
    requirement="The Total row at the bottom must show the correct Net, VAT and Gross totals.",
)
def totals_correct(ctx):
    values = ctx.values()
    t = _total_row(ctx)
    net = sum(_net(line) for line in ctx.meta["lines"])
    rate = ctx.meta["rate"]
    for col, label, exp in ((4, "Net", net), (5, "VAT", net * rate), (6, "Gross", net * (1 + rate))):
        got = ctx.num(values, X.make_ref(SHEET, (t, col)), f"Total {label}")
        require(X.close(got, exp), f"Total {label} is {got:,.2f}, expected {exp:,.2f}")


@rule(
    "invoice-unchanged",
    "Line items, quantities, prices and the rate are unchanged",
    kind="preserve",
    title_zh="明细、数量、单价和税率保持不变",
    requirement="Do not change anything the task does not ask you to change.",
)
def invoice_unchanged(ctx):
    vals = ctx.values().ws(SHEET)
    require(X.close(vals["B1"].value, ctx.meta["rate"]), f"the rate in B1 changed to {X.show(vals['B1'].value)}")
    for r, line in zip(_rows(ctx), ctx.meta["lines"]):
        require(vals.cell(r, 1).value == line["item"], f"row {r} no longer holds '{line['item']}'")
        require(X.close(vals.cell(r, 2).value, line["qty"]), f"Qty of '{line['item']}' changed")
        require(X.close(vals.cell(r, 3).value, line["price"]), f"Unit Price of '{line['item']}' changed")
        require(X.close(vals.cell(r, 4).value, _net(line)), f"Net of '{line['item']}' changed")
    require(X.norm_text(vals.cell(_total_row(ctx), 1).value) == "total", "the Total row moved or was relabelled")
