import openpyxl

from failbook import kit
from failbook import xlsx as X
from failbook.checks import require, rule

SHEET = "Products"
HEADERS = ["SKU", "Product", "Cost", "Price"]


def params(seed):
    r = kit.rng(seed, "xlsx-002")
    n = r.randint(12, 18)
    names = kit.sample(r, kit.PRODUCTS, n)
    products = []
    for i, name in enumerate(names):
        cost = round(r.uniform(2, 80), 2)
        price = round(cost * r.uniform(1.15, 2.4), 2)
        products.append({"sku": f"P-{r.randint(1000, 9999)}{i}", "name": name, "cost": cost, "price": price})
    return {"products": products}


def build(meta, outdir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.append(HEADERS)
    kit.style_header(ws, 1, len(HEADERS))
    for i, p in enumerate(meta["products"], start=2):
        ws.append([p["sku"], p["name"], p["cost"], p["price"]])
        ws.cell(i, 3).number_format = kit.MONEY_FMT
        ws.cell(i, 4).number_format = kit.MONEY_FMT
    kit.set_widths(ws, {"A": 10, "B": 18, "C": 10, "D": 10, "E": 11})
    ws.freeze_panes = "A2"
    wb.save(outdir / "pricing.xlsx")


def _margin_col(ctx):
    ws = ctx.ws(SHEET)
    for cell in ws[1]:
        if cell.value and "margin" in X.norm_text(cell.value):
            return cell.column
    raise AssertionError("no 'Margin %' column header found in row 1")


def _rows(ctx):
    return range(2, 2 + len(ctx.meta["products"]))


def _expected(p, cost=None, price=None):
    cost = p["cost"] if cost is None else cost
    price = p["price"] if price is None else price
    return (price - cost) / price


def scenarios(ctx):
    n = len(ctx.meta["products"])
    first, last = ctx.meta["products"][0], ctx.meta["products"][-1]
    return {
        "reprice": {
            X.make_ref(SHEET, (2, 4)): round(first["price"] * 1.5, 2),
            X.make_ref(SHEET, (n + 1, 3)): round(last["cost"] * 0.5, 2),
        }
    }


@rule(
    "margin-column-added",
    "A 'Margin %' column is added right after Price (column E)",
    kind="content",
    title_zh="在 Price 之后（E 列）新增了 Margin % 列",
    requirement="Add a new column Margin % in column E, right after Price.",
)
def margin_column_added(ctx):
    require(_margin_col(ctx) == 5, f"the Margin % column is in column {_margin_col(ctx)}, expected E")


@rule(
    "margins-correct",
    "Every product's margin equals (Price - Cost) / Price",
    kind="value",
    title_zh="每个产品的毛利率都等于 (Price − Cost) / Price",
    requirement="Fill it in for every product row.",
)
def margins_correct(ctx):
    col = _margin_col(ctx)
    values = ctx.values()
    for r, p in zip(_rows(ctx), ctx.meta["products"]):
        got = ctx.num(values, X.make_ref(SHEET, (r, col)), f"margin of {p['sku']}")
        require(X.close(got, _expected(p)), f"margin of {p['sku']} is {got:.4f}, expected {_expected(p):.4f}")


@rule(
    "margins-are-live",
    "Margins update when a Price or a Cost changes",
    kind="live",
    title_zh="修改 Price 或 Cost 后，毛利率随之更新",
    requirement="It must stay correct when someone later changes a Cost or a Price, so calculate it with formulas.",
)
def margins_are_live(ctx):
    col = _margin_col(ctx)
    values = ctx.values(scenario="reprice")
    products = ctx.meta["products"]
    first, last = products[0], products[-1]
    checks = [
        (2, first, _expected(first, price=round(first["price"] * 1.5, 2)), "its Price"),
        (1 + len(products), last, _expected(last, cost=round(last["cost"] * 0.5, 2)), "its Cost"),
    ]
    for r, p, exp, what in checks:
        got = ctx.num(values, X.make_ref(SHEET, (r, col)), f"margin of {p['sku']}")
        require(
            X.close(got, exp),
            f"after changing {what}, the margin of {p['sku']} is {got:.4f}, expected {exp:.4f}; "
            "the column holds fixed numbers instead of formulas",
        )


@rule(
    "percent-format",
    "The margin cells are formatted as percentages",
    kind="structure",
    title_zh="毛利率单元格为百分比格式",
    requirement="Format the column as a percentage.",
)
def percent_format(ctx):
    col = _margin_col(ctx)
    ws = ctx.ws(SHEET)
    bad = [r for r in _rows(ctx) if not X.is_percent_format(X.number_format(ws, r, col))]
    require(not bad, f"{len(bad)} margin cell(s) are not formatted as a percentage (e.g. row {bad[0] if bad else ''})")


@rule(
    "products-unchanged",
    "SKU, Product, Cost and Price are unchanged",
    kind="preserve",
    title_zh="SKU、Product、Cost、Price 保持不变",
    requirement="Do not change anything the task does not ask you to change.",
)
def products_unchanged(ctx):
    vals = ctx.values().ws(SHEET)
    for r, p in zip(_rows(ctx), ctx.meta["products"]):
        row = [vals.cell(r, c).value for c in range(1, 5)]
        require(row[0] == p["sku"] and row[1] == p["name"], f"row {r} no longer holds {p['sku']} / {p['name']}")
        require(X.close(row[2], p["cost"]) and X.close(row[3], p["price"]), f"Cost or Price of {p['sku']} changed")
