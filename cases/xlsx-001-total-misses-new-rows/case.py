import datetime as dt

import openpyxl

from failbook import kit
from failbook import xlsx as X
from failbook.checks import Fail, require, rule

SHEET = "Orders"
HEADERS = ["Order ID", "Date", "Customer", "Qty", "Unit Price", "Amount"]
PRICES = [12.5, 19.9, 24.0, 35.0, 48.75, 99.0, 120.0, 7.25]


def params(seed):
    r = kit.rng(seed, "xlsx-001")
    n = r.randint(8, 12)
    start = 1001 + r.randint(0, 300)
    day = dt.date(2024, 3, 1)
    orders = []
    for i in range(n + 3):
        day += dt.timedelta(days=r.randint(0, 2))
        orders.append(
            {
                "id": f"SO-{start + i}",
                "date": day.isoformat(),
                "customer": r.choice(kit.COMPANIES),
                "qty": r.randint(1, 40),
                "price": r.choice(PRICES),
            }
        )
    existing, new = orders[:n], orders[n:]
    table = kit.md_table(
        HEADERS[:5], [(o["id"], o["date"], o["customer"], o["qty"], f"{o['price']:.2f}") for o in new]
    )
    return {"existing": existing, "new": new, "new_orders_md": table}


def build(meta, outdir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.append(HEADERS)
    kit.style_header(ws, 1, len(HEADERS))
    for i, o in enumerate(meta["existing"], start=2):
        ws.append([o["id"], kit.from_iso(o["date"]), o["customer"], o["qty"], o["price"], f"=D{i}*E{i}"])
        ws.cell(i, 2).number_format = kit.DATE_FMT
        ws.cell(i, 5).number_format = kit.MONEY_FMT
        ws.cell(i, 6).number_format = kit.MONEY_FMT
    last = len(meta["existing"]) + 1
    total = last + 1
    ws.cell(total, 1, "Total").font = kit.TOTAL_FONT
    ws.cell(total, 4, f"=SUM(D2:D{last})").font = kit.TOTAL_FONT
    ws.cell(total, 6, f"=SUM(F2:F{last})").font = kit.TOTAL_FONT
    ws.cell(total, 6).number_format = kit.MONEY_FMT
    kit.set_widths(ws, {"A": 11, "B": 12, "C": 18, "D": 7, "E": 11, "F": 12})
    ws.freeze_panes = "A2"
    wb.save(outdir / "sales.xlsx")


# -- helpers used by the rules -------------------------------------------------------
def _layout(ctx):
    ws = ctx.ws(SHEET)
    cols = X.header_map(ws, 1)
    missing = [h for h in HEADERS if X.norm_text(h) not in cols]
    require(not missing, f"header(s) missing on {SHEET}: {', '.join(missing)}")
    total = X.find_cells(ws, "Total", col=cols["order id"])
    require(total, "no row labelled 'Total' in the Order ID column")
    return ws, {h: cols[X.norm_text(h)] for h in HEADERS}, total[-1][0]


def _order_rows(ctx):
    ws, cols, total_row = _layout(ctx)
    rows = {}
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, cols["Order ID"]).value
        if isinstance(v, str) and v.strip().upper().startswith("SO-"):
            rows.setdefault(v.strip().upper(), []).append(r)
    return ws, cols, total_row, rows


def _all_orders(meta):
    return meta["existing"] + meta["new"]


def scenarios(ctx):
    ws, cols, total_row, rows = _order_rows(ctx)
    last_new = ctx.meta["new"][-1]["id"]
    require(last_new in rows, f"new order {last_new} not found")
    r = rows[last_new][0]
    return {"bump": {X.make_ref(SHEET, (r, cols["Qty"])): ctx.meta["new"][-1]["qty"] + 10}}


# -- rules -------------------------------------------------------------------------------
@rule(
    "new-orders-added",
    "Each new order appears once, with the given date, customer, qty and price",
    kind="content",
    title_zh="每笔新订单出现且仅出现一次，日期、客户、数量和单价正确",
    requirement="Add the following three new orders to the log.",
)
def new_orders_added(ctx):
    _, cols, _, rows = _order_rows(ctx)
    vals = ctx.values().ws(SHEET)
    for o in ctx.meta["new"]:
        hits = rows.get(o["id"], [])
        require(len(hits) == 1, f"order {o['id']} appears {len(hits)} times")
        r = hits[0]
        date = vals.cell(r, cols["Date"]).value
        require(X.is_date(date), f"{o['id']}: Date is {X.show(date)}, not a real date")
        require(X.as_date(date) == kit.from_iso(o["date"]), f"{o['id']}: Date is {X.show(date)}, expected {o['date']}")
        require(X.norm_text(vals.cell(r, cols["Customer"]).value) == X.norm_text(o["customer"]),
                f"{o['id']}: Customer is {X.show(vals.cell(r, cols['Customer']).value)}")
        for key, col in (("qty", "Qty"), ("price", "Unit Price")):
            got = vals.cell(r, cols[col]).value
            require(X.is_number(got), f"{o['id']}: {col} is {X.show(got)}, not a number")
            require(X.close(got, o[key]), f"{o['id']}: {col} is {X.show(got)}, expected {o[key]}")


@rule(
    "existing-orders-unchanged",
    "The original orders are unchanged and in their original order",
    kind="preserve",
    title_zh="原有订单保持不变，顺序不变",
    requirement="Do not change anything the task does not ask you to change.",
)
def existing_orders_unchanged(ctx):
    _, cols, _, rows = _order_rows(ctx)
    vals = ctx.values().ws(SHEET)
    previous = 0
    for o in ctx.meta["existing"]:
        hits = rows.get(o["id"], [])
        require(len(hits) == 1, f"original order {o['id']} appears {len(hits)} times")
        r = hits[0]
        require(r > previous, f"original order {o['id']} moved out of order")
        previous = r
        require(X.close(vals.cell(r, cols["Qty"]).value, o["qty"]), f"{o['id']}: Qty changed")
        require(X.close(vals.cell(r, cols["Unit Price"]).value, o["price"]), f"{o['id']}: Unit Price changed")
        require(X.as_date(vals.cell(r, cols["Date"]).value) == kit.from_iso(o["date"]), f"{o['id']}: Date changed")


@rule(
    "amounts-correct",
    "Amount equals Qty x Unit Price on every order row",
    kind="value",
    title_zh="每一行的 Amount 都等于 Qty × Unit Price",
    requirement="The Amount of each new order is Qty x Unit Price, calculated the same way as in the existing rows.",
)
def amounts_correct(ctx):
    _, cols, _, rows = _order_rows(ctx)
    values = ctx.values()
    for o in _all_orders(ctx.meta):
        for r in rows.get(o["id"], []):
            got = ctx.num(values, X.make_ref(SHEET, (r, cols["Amount"])), f"Amount of {o['id']}")
            require(X.close(got, o["qty"] * o["price"]), f"Amount of {o['id']} is {got:.2f}, expected {o['qty'] * o['price']:.2f}")


@rule(
    "totals-include-all-orders",
    "The Total row sums Qty and Amount over all orders, old and new",
    kind="value",
    title_zh="合计行的 Qty 与 Amount 覆盖全部订单（新旧都算）",
    requirement="The Total row must include all orders, including the new ones.",
)
def totals_include_all_orders(ctx):
    _, cols, total_row, _ = _order_rows(ctx)
    values = ctx.values()
    orders = _all_orders(ctx.meta)
    exp_qty = sum(o["qty"] for o in orders)
    exp_amt = sum(o["qty"] * o["price"] for o in orders)
    got_qty = ctx.num(values, X.make_ref(SHEET, (total_row, cols["Qty"])), "Total Qty")
    got_amt = ctx.num(values, X.make_ref(SHEET, (total_row, cols["Amount"])), "Total Amount")
    require(X.close(got_qty, exp_qty), f"Total Qty is {got_qty:g}, expected {exp_qty} (difference {exp_qty - got_qty:g})")
    require(X.close(got_amt, exp_amt), f"Total Amount is {got_amt:,.2f}, expected {exp_amt:,.2f} (short by {exp_amt - got_amt:,.2f})")


@rule(
    "totals-are-live",
    "Changing a new order's Qty updates its Amount and the totals",
    kind="live",
    title_zh="修改新订单的数量后，金额与合计随之更新",
    requirement="Keep the totals as formulas so that they update when the data changes.",
)
def totals_are_live(ctx):
    _, cols, total_row, _ = _order_rows(ctx)
    base, bumped = ctx.values(), ctx.values(scenario="bump")
    price = ctx.meta["new"][-1]["price"]
    q = X.make_ref(SHEET, (total_row, cols["Qty"]))
    a = X.make_ref(SHEET, (total_row, cols["Amount"]))
    dq = ctx.num(bumped, q, "Total Qty") - ctx.num(base, q, "Total Qty")
    da = ctx.num(bumped, a, "Total Amount") - ctx.num(base, a, "Total Amount")
    who = ctx.meta["new"][-1]["id"]
    require(X.close(dq, 10), f"adding 10 to the Qty of {who} changed Total Qty by {dq:g}, expected +10")
    require(X.close(da, 10 * price), f"adding 10 to the Qty of {who} changed Total Amount by {da:,.2f}, expected +{10 * price:,.2f}")


@rule(
    "total-row-last",
    "The Total row is directly below the last order",
    kind="content",
    title_zh="合计行紧接在最后一笔订单之后",
    requirement="Put them after the existing orders and directly above the Total row. The Total row stays the last row of the table.",
)
def total_row_last(ctx):
    _, _, total_row, rows = _order_rows(ctx)
    last_order = max(r for hits in rows.values() for r in hits)
    require(last_order < total_row, f"an order is below the Total row (row {last_order} > {total_row})")
    require(total_row == last_order + 1, f"there is a gap between the last order (row {last_order}) and the Total row ({total_row})")
