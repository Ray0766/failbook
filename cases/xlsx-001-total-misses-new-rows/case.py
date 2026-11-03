import datetime as dt

import openpyxl

from failbook import kit
from failbook import xlsx as X
from failbook.checks import Fail, Inconclusive, require, rule

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


def _amt(x, sign=False):
    """A money amount with as many decimals as it needs (at least 2, at most 6)."""
    whole, _, frac = (f"{x:+,.6f}" if sign else f"{x:,.6f}").rstrip("0").partition(".")
    return f"{whole}.{frac.ljust(2, '0')}"


#: What X.number_format_key counts as one format, as the format messages say it.
_SAME_FORMAT = ("(formats count as the same when they differ only in letter case, quotes or backslash escapes, a "
                "text section that is only '@', a zero section equal to the first, a negative section of '-' and a "
                "first section of digits such as #,##0.00;-#,##0.00 (it shows -0.00 instead of 0.00 only for a "
                "negative number that rounds to zero, and the task's prices and amounts are all positive), or a "
                "locale tag such as [$-409] on a date of digits; under such a tag an unquoted '/' counts as the "
                "date separator LibreOffice 24.2 shows for it: [$-41D]yyyy/mm/dd shows 2024-03-02, "
                "[$-409]yyyy/mm/dd 2024/03/02 and [$-407]yyyy/mm/dd 2024.03.02).")


#: The "bump" scenario changes, in a copy of the file, every order's Qty (by a different
#: amount per order) and every Unit Price, so a typed Amount or a total that leaves out
#: any row no longer matches.
PRICE_BUMP = 1.0


def _bumped(ctx):
    """[(order, row, new qty, new price)] for every order row found in the output."""
    _, _, _, rows = _order_rows(ctx)
    out = []
    for k, o in enumerate(_all_orders(ctx.meta), start=1):
        for r in rows.get(o["id"], []):
            out.append((o, r, o["qty"] + k, o["price"] + PRICE_BUMP))
    return out


def scenarios(ctx):
    _, cols, _, _ = _order_rows(ctx)
    overrides = {}
    for _, r, qty, price in _bumped(ctx):
        overrides[X.make_ref(SHEET, (r, cols["Qty"]))] = qty
        overrides[X.make_ref(SHEET, (r, cols["Unit Price"]))] = price
    return {"bump": overrides}


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
    "new-rows-formatted",
    "The new orders' Date, Unit Price and Amount have the number formats of the original order rows",
    kind="structure",
    title_zh="新订单的 Date、Unit Price 和 Amount 与原有订单行的数字格式相同",
    requirement="Put them after the existing orders and directly above the Total row, in the same format as the "
                "other order rows.",
)
def new_rows_formatted(ctx):
    ws, cols, _, rows = _order_rows(ctx)
    given = ctx.input_wb()[SHEET]
    # In the file as given the original orders are rows 2 onwards, all formatted alike. Columns
    # they leave in General (Order ID, Customer, Qty) are not compared.
    general = X.number_format_key("General")
    wanted = {h: X.number_format(given, 2, c) for c, h in enumerate(HEADERS, start=1)}
    wanted = {h: f for h, f in wanted.items() if X.number_format_key(f) != general}
    found = [(o, r) for o in ctx.meta["new"] for r in rows.get(o["id"], [])]
    require(found, "none of the new orders is in the sheet, so no new row's format can be compared")
    wrong, unsure, why = [], [], {}  # why: column -> why the checker cannot compare it
    for o, r in found:
        diff, maybe = [], []
        for h, f in wanted.items():
            got = X.number_format(ws, r, cols[h])
            if X.number_format_key(got) != X.number_format_key(f):
                reason = X.number_format_unsure(got, f)  # '/' under a locale tag: the separator is unknown
                (maybe if reason else diff).append(f"{h} '{got}'")
                if reason:
                    why[h] = reason
        if diff:
            wrong.append(f"{o['id']} (row {r}) has " + ", ".join(diff))
        elif maybe:
            unsure.append(f"{o['id']} (row {r}) has " + ", ".join(maybe))
    given_text = ", ".join(f"{h} '{f}'" for h, f in wanted.items())
    require(not wrong, f"{len(wrong)} of {len(found)} new order rows are not formatted like the original order rows: "
                       + "; ".join(wrong[:3]) + f". In {ctx.case.inputs[0]} as given, the original order rows have "
                       f"{given_text} {_SAME_FORMAT}")
    if unsure:
        reason = next(iter(why.values()))
        raise Inconclusive(f"{len(unsure)} of {len(found)} new order rows have a format that matches the original "
                           "order rows' except for '/' under a locale tag: " + "; ".join(unsure[:3])
                           + f". {reason[0].upper()}{reason[1:]}, so it cannot tell whether these cells show like "
                           + ", ".join(f"{h} '{wanted[h]}'" for h in why) + " in the original order rows")


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
            exp = o["qty"] * o["price"]
            require(X.close(got, exp), f"Amount of {o['id']} is {_amt(got)}, expected {_amt(exp)} "
                                      f"(difference {_amt(got - exp, sign=True)})")


@rule(
    "amounts-are-live",
    "Every order's Amount follows its Qty and Unit Price (new rows and original rows)",
    kind="live",
    title_zh="每笔订单（新行和原有行）的 Amount 都随 Qty 和 Unit Price 变化",
    requirement="The Amount of each new order is Qty x Unit Price, calculated the same way as in the existing rows "
                "(with a formula). Do not change anything the task does not ask you to change.",
)
def amounts_are_live(ctx):
    _, cols, _, _ = _order_rows(ctx)
    base, bumped = ctx.values(), ctx.values(scenario="bump")
    new_ids = {o["id"] for o in ctx.meta["new"]}
    changes = _bumped(ctx)
    wrong = []
    for o, r, qty, price in changes:
        ref = X.make_ref(SHEET, (r, cols["Amount"]))
        got = ctx.num(bumped, ref, f"Amount of {o['id']}")
        exp = qty * price
        if X.close(got, exp):
            continue
        which = "new order" if o["id"] in new_ids else "original order"
        note = ""
        before = base[ref]
        if X.is_number(before) and X.close(got, before):
            note = ", the same as before the change"
        wrong.append(f"{which} {o['id']} (row {r}): with Qty {qty} and Unit Price {price:.2f} its Amount is "
                     f"{_amt(got)}, expected {_amt(exp)}{note}")
    require(not wrong, f"in a copy of the file where every order's Qty and Unit Price were changed, {len(wrong)} of "
                       f"{len(changes)} Amounts did not follow: " + "; ".join(wrong[:3]))


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
    require(X.close(got_qty, exp_qty), f"Total Qty is {got_qty:g}, expected {exp_qty} (difference {got_qty - exp_qty:+g})")
    require(X.close(got_amt, exp_amt), f"Total Amount is {_amt(got_amt)}, expected {_amt(exp_amt)} "
                                       f"(difference {_amt(got_amt - exp_amt, sign=True)})")


@rule(
    "totals-are-live",
    "Changing the Qty and Unit Price of the orders updates Total Qty and Total Amount",
    kind="live",
    title_zh="修改各订单的数量和单价后，合计的 Qty 与 Amount 随之更新",
    requirement="Keep the totals as formulas so that they update when the data changes.",
)
def totals_are_live(ctx):
    _, cols, total_row, _ = _order_rows(ctx)
    base, bumped = ctx.values(), ctx.values(scenario="bump")
    changes = _bumped(ctx)
    exp_dq = sum(qty - o["qty"] for o, _, qty, _ in changes)
    exp_da = sum(qty * price - o["qty"] * o["price"] for o, _, qty, price in changes)
    q = X.make_ref(SHEET, (total_row, cols["Qty"]))
    a = X.make_ref(SHEET, (total_row, cols["Amount"]))
    dq = ctx.num(bumped, q, "Total Qty") - ctx.num(base, q, "Total Qty")
    da = ctx.num(bumped, a, "Total Amount") - ctx.num(base, a, "Total Amount")
    what = (f"in a copy of the file where each order's Qty was raised by a different amount ({exp_dq:+g} in all) "
            f"and every Unit Price by {PRICE_BUMP:.2f}")
    require(X.close(dq, exp_dq), f"{what}, Total Qty changed by {dq:+g}, expected {exp_dq:+g} (difference {dq - exp_dq:+g})")
    if not X.close(da, exp_da):
        shown = [bumped[X.make_ref(SHEET, (r, cols["Amount"]))] for _, r, _, _ in changes]
        note = ""
        if all(X.is_number(v) for v in shown) and X.close(sum(shown), ctx.num(bumped, a, "Total Amount")):
            note = ("; Total Amount does equal the sum of the order Amounts in that copy, so the gap comes from "
                    "Amounts that did not follow (see amounts-are-live)")
        raise Fail(f"{what}, Total Amount changed by {_amt(da, sign=True)}, expected {_amt(exp_da, sign=True)} "
                   f"(difference {_amt(da - exp_da, sign=True)}){note}")


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
