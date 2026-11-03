import openpyxl

from failbook import kit
from failbook import xlsx as X
from failbook.checks import require, rule

PRODUCTS, ORDERS = "Products", "Orders"
CATEGORIES = ["Furniture", "Stationery", "Electronics", "Kitchen", "Storage"]


def params(seed):
    r = kit.rng(seed, "xlsx-009")
    n = r.randint(12, 16)
    skus = r.sample(range(100, 999), n)
    products = [
        {"sku": f"P-{s}", "category": r.choice(CATEGORIES), "price": round(r.uniform(3, 400), 2)} for s in skus
    ]
    # The list must not be sorted by SKU, otherwise approximate match would happen to work.
    while products == sorted(products, key=lambda p: p["sku"]):
        r.shuffle(products)
    orders = []
    for i in range(r.randint(18, 24)):
        p = r.choice(products)
        orders.append({"order": f"O-{5001 + i}", "sku": p["sku"], "qty": r.randint(1, 30)})
    return {"products": products, "orders": orders}


def build(meta, outdir):
    wb = openpyxl.Workbook()
    o = wb.active
    o.title = ORDERS
    o.append(["Order", "SKU", "Qty", "Category", "Unit Price", "Amount"])
    kit.style_header(o, 1, 6)
    for order in meta["orders"]:
        o.append([order["order"], order["sku"], order["qty"]])
    kit.set_widths(o, {"A": 10, "B": 9, "C": 7, "D": 13, "E": 11, "F": 12})
    p = wb.create_sheet(PRODUCTS)
    p.append(["SKU", "Category", "Unit Price"])
    kit.style_header(p, 1, 3)
    for i, prod in enumerate(meta["products"], start=2):
        p.append([prod["sku"], prod["category"], prod["price"]])
        p.cell(i, 3).number_format = kit.MONEY_FMT
    kit.set_widths(p, {"A": 9, "B": 13, "C": 11})
    wb.save(outdir / "orders.xlsx")


def _product(meta, sku):
    return next(p for p in meta["products"] if p["sku"] == sku)


def _rows(ctx):
    return list(zip(range(2, 2 + len(ctx.meta["orders"])), ctx.meta["orders"]))


def scenarios(ctx):
    ws = ctx.ws(PRODUCTS)
    sku = ctx.meta["orders"][0]["sku"]
    hit = X.find_cell(ws, sku, col=1)
    require(hit, f"product {sku} not found on the Products sheet")
    p = _product(ctx.meta, sku)
    return {"reprice": {X.make_ref(PRODUCTS, (hit[0], 3)): round(p["price"] + 10, 2)}}


def _mismatches(ctx, col, expected):
    values = ctx.values()
    bad = []
    for r, order in _rows(ctx):
        ref = X.make_ref(ORDERS, (r, col))
        exp = expected(order)
        if isinstance(exp, str):
            got = ctx.value(values, ref)
            ok = X.norm_text(got) == X.norm_text(exp)
        else:
            got = ctx.num(values, ref, f"{order['order']}")
            ok = X.close(got, exp)
        if not ok:
            bad.append(f"{order['order']} ({order['sku']}): {X.show(got)} instead of {X.show(exp)}")
    return bad


@rule(
    "categories-correct",
    "Every order shows the category of its own SKU",
    kind="value",
    title_zh="每个订单显示的都是其 SKU 对应的类别",
    requirement="Fill in Category ... by looking up the order's SKU in the Products sheet.",
)
def categories_correct(ctx):
    bad = _mismatches(ctx, 4, lambda o: _product(ctx.meta, o["sku"])["category"])
    require(not bad, f"{len(bad)} of {len(ctx.meta['orders'])} categories are wrong, e.g. " + "; ".join(bad[:3]))


@rule(
    "prices-correct",
    "Every order shows the unit price of its own SKU",
    kind="value",
    title_zh="每个订单显示的都是其 SKU 对应的单价",
    requirement="Fill in ... Unit Price by looking up the order's SKU in the Products sheet.",
)
def prices_correct(ctx):
    bad = _mismatches(ctx, 5, lambda o: _product(ctx.meta, o["sku"])["price"])
    require(not bad, f"{len(bad)} of {len(ctx.meta['orders'])} unit prices are wrong, e.g. " + "; ".join(bad[:3]))


@rule(
    "amounts-correct",
    "Amount = Qty x the product's unit price on every order",
    kind="value",
    title_zh="每个订单 Amount = Qty × 该产品单价",
    requirement="Fill in Amount = Qty x Unit Price.",
)
def amounts_correct(ctx):
    bad = _mismatches(ctx, 6, lambda o: o["qty"] * _product(ctx.meta, o["sku"])["price"])
    require(not bad, f"{len(bad)} of {len(ctx.meta['orders'])} amounts are wrong, e.g. " + "; ".join(bad[:3]))


@rule(
    "lookups-live",
    "Changing a price on Products updates the matching orders",
    kind="live",
    title_zh="修改 Products 表中的价格后，对应订单随之更新",
    requirement="Use formulas, so that the orders update automatically if a product's category or price is changed.",
)
def lookups_live(ctx):
    values = ctx.values(scenario="reprice")
    sku = ctx.meta["orders"][0]["sku"]
    new_price = round(_product(ctx.meta, sku)["price"] + 10, 2)
    for r, order in _rows(ctx):
        if order["sku"] != sku:
            continue
        got = ctx.num(values, X.make_ref(ORDERS, (r, 5)), f"Unit Price of {order['order']}")
        require(X.close(got, new_price), f"after changing the price of {sku} to {new_price}, {order['order']} shows {got}")
        amt = ctx.num(values, X.make_ref(ORDERS, (r, 6)), f"Amount of {order['order']}")
        require(X.close(amt, new_price * order["qty"]), f"after the price change, {order['order']}'s Amount is {amt}")


@rule(
    "data-unchanged",
    "Orders (Order, SKU, Qty) and the Products sheet are unchanged",
    kind="preserve",
    title_zh="订单（Order、SKU、Qty）与 Products 表保持不变",
    requirement="Do not reorder or edit the Products sheet.",
)
def data_unchanged(ctx):
    vals = ctx.values()
    p = vals.ws(PRODUCTS)
    for i, prod in enumerate(ctx.meta["products"], start=2):
        got = (p.cell(i, 1).value, p.cell(i, 2).value, p.cell(i, 3).value)
        require(got[0] == prod["sku"] and got[1] == prod["category"] and X.close(got[2], prod["price"]),
                f"Products row {i} changed: {got}")
    o = vals.ws(ORDERS)
    for r, order in _rows(ctx):
        got = (o.cell(r, 1).value, o.cell(r, 2).value, o.cell(r, 3).value)
        require(got == (order["order"], order["sku"], order["qty"]), f"Orders row {r} changed: {got}")
