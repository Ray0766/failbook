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


def _repriced(p):
    """Cost and Price of a product in the "reprice" scenario (a copy of the file in which both change)."""
    return round(p["cost"] * 0.5, 2), round(p["price"] * 1.5, 2)


def scenarios(ctx):
    overrides = {}
    for r, p in zip(_rows(ctx), ctx.meta["products"]):
        cost, price = _repriced(p)
        overrides[X.make_ref(SHEET, (r, 3))] = cost
        overrides[X.make_ref(SHEET, (r, 4))] = price
    return {"reprice": overrides}


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
    requirement="Margin % = (Price - Cost) / Price. Fill it in for every product row. Do not round the margin.",
)
def margins_correct(ctx):
    col = _margin_col(ctx)
    values = ctx.values()
    wrong = []
    for r, p in zip(_rows(ctx), ctx.meta["products"]):
        got = ctx.num(values, X.make_ref(SHEET, (r, col)), f"margin of {p['sku']}")
        exp = _expected(p)
        if not X.close(got, exp):
            wrong.append(f"{p['sku']} (row {r}): {got:.6g}, expected {exp:.6g} (difference {got - exp:+.2g})")
    require(not wrong, f"{len(wrong)} of {len(ctx.meta['products'])} margins are wrong: " + "; ".join(wrong[:3]))


@rule(
    "margins-are-live",
    "Margins update when a Price or a Cost changes",
    kind="live",
    title_zh="修改 Price 或 Cost 后，毛利率随之更新",
    requirement="It must stay correct when someone later changes a Cost or a Price, so calculate it with formulas.",
)
def margins_are_live(ctx):
    col = _margin_col(ctx)
    base, values = ctx.values(), ctx.values(scenario="reprice")
    wrong = []
    for r, p in zip(_rows(ctx), ctx.meta["products"]):
        ref = X.make_ref(SHEET, (r, col))
        got = ctx.num(values, ref, f"margin of {p['sku']}")
        cost, price = _repriced(p)
        exp = _expected(p, cost=cost, price=price)
        if X.close(got, exp):
            continue
        before = base[ref]
        same = ", the same as before the change" if X.is_number(before) and X.close(before, got) else ""
        wrong.append(f"{p['sku']} (row {r}): with Cost {cost:.2f} and Price {price:.2f} its margin is {got:.6g}, "
                     f"expected {exp:.6g} (difference {got - exp:+.2g}){same}")
    require(not wrong, f"in a copy of the file where every Cost and Price was changed, {len(wrong)} of "
                       f"{len(ctx.meta['products'])} margins did not follow: " + "; ".join(wrong[:3]))


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
