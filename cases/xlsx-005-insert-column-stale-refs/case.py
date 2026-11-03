import openpyxl

from failbook import kit
from failbook import xlsx as X
from failbook.checks import require, rule

LINES, SUMMARY = "Lines", "Summary"
OLD_HEADERS = ["Product", "Unit Price", "Qty", "Line Total"]
NEW_HEADERS = ["Product", "Unit Price", "Discount", "Qty", "Line Total"]
LABELS = {"revenue": "Total revenue", "units": "Total units", "avg": "Average unit price"}


def params(seed):
    r = kit.rng(seed, "xlsx-005")
    n = r.randint(8, 12)
    products = kit.sample(r, kit.PRODUCTS, n)
    lines = [{"product": p, "price": round(r.uniform(4, 180), 2), "qty": r.randint(1, 60)} for p in products]
    discounted = r.sample(products, r.randint(3, 5))
    discounts = [{"product": p, "discount": r.choice([0.05, 0.1, 0.15, 0.2, 0.25])} for p in discounted]
    table = kit.md_table(["Product", "Discount"], [(d["product"], f"{d['discount']:.0%}") for d in discounts])
    return {"lines": lines, "discounts": discounts, "discounts_md": table}


def build(meta, outdir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = LINES
    ws.append(OLD_HEADERS)
    kit.style_header(ws, 1, len(OLD_HEADERS))
    for r, line in enumerate(meta["lines"], start=2):
        ws.append([line["product"], line["price"], line["qty"], f"=B{r}*C{r}"])
        ws.cell(r, 2).number_format = kit.MONEY_FMT
        ws.cell(r, 4).number_format = kit.MONEY_FMT
    kit.set_widths(ws, {"A": 18, "B": 11, "C": 10, "D": 12, "E": 12})
    last = 1 + len(meta["lines"])
    s = wb.create_sheet(SUMMARY)
    s["A1"], s["B1"] = LABELS["revenue"], f"=SUM(Lines!D2:D{last})"
    s["A2"], s["B2"] = LABELS["units"], f"=SUM(Lines!C2:C{last})"
    s["A3"], s["B3"] = LABELS["avg"], f"=AVERAGE(Lines!B2:B{last})"
    for cell in ("B1", "B3"):
        s[cell].number_format = kit.MONEY_FMT
    for row in range(1, 4):
        s.cell(row, 1).font = kit.TOTAL_FONT
    kit.set_widths(s, {"A": 20, "B": 14})
    wb.save(outdir / "orders.xlsx")


def _discount(meta, product):
    for d in meta["discounts"]:
        if d["product"] == product:
            return d["discount"]
    return 0.0


def _total(meta, line, qty=None):
    qty = line["qty"] if qty is None else qty
    return line["price"] * qty * (1 - _discount(meta, line["product"]))


def _rows(ctx):
    return range(2, 2 + len(ctx.meta["lines"]))


def _summary_cell(ctx, key):
    ws = ctx.ws(SUMMARY)
    hit = X.find_cell(ws, LABELS[key], col=1)
    require(hit, f"label {LABELS[key]!r} not found on the Summary sheet")
    return X.make_ref(SUMMARY, (hit[0], 2))


def _layout_ok(ctx):
    ws = ctx.ws(LINES)
    got = [ws.cell(1, c).value for c in range(1, 6)]
    require([X.norm_text(v) for v in got] == [X.norm_text(h) for h in NEW_HEADERS],
            f"Lines headers are {got}, expected {NEW_HEADERS}")


def scenarios(ctx):
    _layout_ok(ctx)
    return {
        "more_units": {X.make_ref(LINES, (2, 4)): ctx.meta["lines"][0]["qty"] + 5},
        "half_off": {X.make_ref(LINES, (2, 3)): 0.5},
    }


@rule(
    "columns-in-order",
    "Lines columns are Product, Unit Price, Discount, Qty, Line Total",
    kind="content",
    title_zh="Lines 表的列依次为 Product、Unit Price、Discount、Qty、Line Total",
    requirement="Insert a new column Discount between Unit Price and Qty.",
)
def columns_in_order(ctx):
    _layout_ok(ctx)


@rule(
    "discounts-entered",
    "Each line has the listed discount as a fraction (0 when not listed)",
    kind="content",
    title_zh="每行的折扣按清单以小数填写（未列出的为 0）",
    requirement="Enter these discounts (as fractions, e.g. 5% is 0.05). Products that are not listed get a discount of 0.",
)
def discounts_entered(ctx):
    _layout_ok(ctx)
    vals = ctx.values().ws(LINES)
    for r, line in zip(_rows(ctx), ctx.meta["lines"]):
        got = vals.cell(r, 3).value
        exp = _discount(ctx.meta, line["product"])
        if got is None and exp == 0:
            continue
        require(X.is_number(got), f"Discount of {line['product']} is {X.show(got)}, not a number")
        hint = " (entered as a whole number instead of a fraction)" if X.close(got, exp * 100) and exp else ""
        require(X.close(got, exp), f"Discount of {line['product']} is {X.show(got)}, expected {exp}{hint}")


@rule(
    "line-totals-correct",
    "Line Total = Unit Price x Qty x (1 - Discount) on every line",
    kind="value",
    title_zh="每行 Line Total = Unit Price × Qty × (1 − Discount)",
    requirement="Change Line Total to Unit Price x Qty x (1 - Discount) ... on every line.",
)
def line_totals_correct(ctx):
    values = ctx.values()
    for r, line in zip(_rows(ctx), ctx.meta["lines"]):
        got = ctx.num(values, X.make_ref(LINES, (r, 5)), f"Line Total of {line['product']}")
        exp = _total(ctx.meta, line)
        require(X.close(got, exp), f"Line Total of {line['product']} is {got:,.2f}, expected {exp:,.2f}")


@rule(
    "line-totals-live",
    "Line Total follows a change of Discount",
    kind="live",
    title_zh="修改 Discount 后，Line Total 随之变化",
    requirement="Change Line Total to Unit Price x Qty x (1 - Discount), as a formula on every line.",
)
def line_totals_live(ctx):
    half = ctx.values(scenario="half_off")
    first = ctx.meta["lines"][0]
    got = ctx.num(half, X.make_ref(LINES, (2, 5)), "Line Total of the first line")
    exp = first["price"] * first["qty"] * 0.5
    require(X.close(got, exp), f"with a 50% discount on the first line, its Line Total is {got:,.2f}, expected {exp:,.2f}")


@rule(
    "summary-correct",
    "Summary shows the correct total revenue, total units and average unit price",
    kind="value",
    title_zh="Summary 显示正确的总收入、总数量和平均单价",
    requirement="The Summary sheet must keep showing the correct total revenue (after discounts), total units and average unit price.",
)
def summary_correct(ctx):
    values = ctx.values()
    lines = ctx.meta["lines"]
    expected = {
        "revenue": sum(_total(ctx.meta, line) for line in lines),
        "units": sum(line["qty"] for line in lines),
        "avg": sum(line["price"] for line in lines) / len(lines),
    }
    for key, exp in expected.items():
        got = ctx.num(values, _summary_cell(ctx, key), LABELS[key])
        require(X.close(got, exp), f"{LABELS[key]} is {got:,.2f}, expected {exp:,.2f}")


@rule(
    "summary-live",
    "Summary updates when a quantity on Lines changes",
    kind="live",
    title_zh="Lines 表的数量变化时，Summary 随之更新",
    requirement="...and keep updating when the Lines data changes.",
)
def summary_live(ctx):
    base, more = ctx.values(), ctx.values(scenario="more_units")
    first = ctx.meta["lines"][0]
    d_units = ctx.num(more, _summary_cell(ctx, "units")) - ctx.num(base, _summary_cell(ctx, "units"))
    d_rev = ctx.num(more, _summary_cell(ctx, "revenue")) - ctx.num(base, _summary_cell(ctx, "revenue"))
    exp_rev = _total(ctx.meta, first, first["qty"] + 5) - _total(ctx.meta, first)
    require(X.close(d_units, 5), f"adding 5 units to the first line changed Total units by {d_units:g}, expected +5")
    require(X.close(d_rev, exp_rev), f"adding 5 units to the first line changed Total revenue by {d_rev:,.2f}, expected +{exp_rev:,.2f}")


@rule(
    "lines-unchanged",
    "Products, unit prices and quantities are unchanged",
    kind="preserve",
    title_zh="产品、单价和数量保持不变",
    requirement="Do not change anything the task does not ask you to change.",
)
def lines_unchanged(ctx):
    _layout_ok(ctx)
    vals = ctx.values().ws(LINES)
    for r, line in zip(_rows(ctx), ctx.meta["lines"]):
        require(vals.cell(r, 1).value == line["product"], f"row {r} no longer holds {line['product']}")
        require(X.close(vals.cell(r, 2).value, line["price"]), f"Unit Price of {line['product']} changed")
        require(X.close(vals.cell(r, 4).value, line["qty"]), f"Qty of {line['product']} changed")
