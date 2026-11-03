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


def _total(meta, line):
    return line["price"] * line["qty"] * (1 - _discount(meta, line["product"]))


def _amt(x, places=2, sign=False):
    """A number with at least ``places`` and at most 6 decimals, so that two numbers the checker
    tells apart never print the same."""
    whole, _, frac = (f"{x:+,.6f}" if sign else f"{x:,.6f}").rstrip("0").partition(".")
    frac = frac.ljust(places, "0")
    return f"{whole}.{frac}" if frac else whole


def _whole_cents(x):
    return abs(x * 100 - round(x * 100)) < 1e-6


def _money_ok(got, exact):
    """The exact amount, or that amount rounded to whole cents (in either direction).

    The task gives Line Total = Unit Price x Qty x (1 - Discount) and does not ask for
    unrounded values; rounding a money amount to the cent is normal, so it is accepted.
    "Within one cent" keeps a round-down of a value whose binary form is just below a whole
    cent from failing.
    """
    return X.close(got, exact) or (_whole_cents(got) and abs(got - exact) <= 0.01 + 1e-9)


def _revenue_ok(values, got, exact, line_exacts):
    """Total revenue is right if it is the exact amount, that amount rounded to whole cents, or the
    sum of the Line Totals shown in the same file when each of them is accepted by _money_ok."""
    if _money_ok(got, exact):
        return True
    shown = [values[X.make_ref(LINES, (r, 5))] for r, _ in line_exacts]
    return (all(X.is_number(v) and _money_ok(v, e) for v, (_, e) in zip(shown, line_exacts))
            and X.close(got, sum(shown)))


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


def _changed(meta):
    """[(row, line, unit price, discount, qty)] in the "changed" scenario.

    In a copy of the file, every line gets a different Unit Price, Discount and Qty
    (each changed by a different amount per line), so a typed Line Total, a formula
    that reads another line, or a Summary figure that is a typed number no longer matches.
    """
    out = []
    for k, line in enumerate(meta["lines"], start=1):
        out.append((k + 1, line, round(line["price"] + 0.5 * k, 2), round(0.3 + 0.01 * k, 2), line["qty"] + k))
    return out


def scenarios(ctx):
    _layout_ok(ctx)
    overrides = {}
    for r, _, price, discount, qty in _changed(ctx.meta):
        overrides[X.make_ref(LINES, (r, 2))] = price
        overrides[X.make_ref(LINES, (r, 3))] = discount
        overrides[X.make_ref(LINES, (r, 4))] = qty
    return {"changed": overrides}


CHANGED = "in a copy of the file where every line's Unit Price, Discount and Qty were changed"


def _same_as_before(base, ref, got):
    before = base[ref]
    return ", the same as before the change" if X.is_number(before) and X.close(before, got) else ""


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
    "discounts-percent-format",
    "The discount cells are formatted as percentages",
    kind="structure",
    title_zh="折扣单元格为百分比格式",
    requirement="Enter these discounts (as fractions, e.g. 5% is 0.05, formatted as a percentage).",
)
def discounts_percent_format(ctx):
    _layout_ok(ctx)
    ws = ctx.ws(LINES)
    bad = []
    for r, line in zip(_rows(ctx), ctx.meta["lines"]):
        if ws.cell(r, 3).value is None:
            continue  # an empty cell shows nothing; discounts-entered decides whether it may be empty
        fmt = X.number_format(ws, r, 3)
        if not X.is_percent_format(fmt):
            bad.append(f"{line['product']} ({X.make_ref(LINES, (r, 3))}, format {fmt!r})")
    require(not bad, f"{len(bad)} discount cell(s) are not formatted as a percentage: " + ", ".join(bad[:3])
            + (f" (and {len(bad) - 3} more)" if len(bad) > 3 else ""))


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
        require(_money_ok(got, exp), f"Line Total of {line['product']} is {_amt(got)}, expected {_amt(exp)} "
                                  f"(difference {_amt(got - exp, sign=True)})")


@rule(
    "line-totals-live",
    "Every Line Total follows changes of its Unit Price, Discount and Qty",
    kind="live",
    title_zh="每行 Line Total 都随本行的 Unit Price、Discount 和 Qty 变化",
    requirement="Change Line Total to Unit Price x Qty x (1 - Discount), as a formula on every line.",
)
def line_totals_live(ctx):
    base, changed = ctx.values(), ctx.values(scenario="changed")
    rows = _changed(ctx.meta)
    wrong = []
    for r, line, price, discount, qty in rows:
        ref = X.make_ref(LINES, (r, 5))
        got = ctx.num(changed, ref, f"Line Total of {line['product']}")
        exp = price * qty * (1 - discount)
        if not _money_ok(got, exp):
            wrong.append(f"{line['product']} (row {r}): with Unit Price {price:.2f}, Discount {discount:.0%} and "
                         f"Qty {qty} its Line Total is {_amt(got)}, expected {_amt(exp)}{_same_as_before(base, ref, got)}")
    require(not wrong, f"{CHANGED}, {len(wrong)} of {len(rows)} Line Totals did not follow: " + "; ".join(wrong[:3]))


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
    line_exacts = [(r, _total(ctx.meta, line)) for r, line in zip(_rows(ctx), lines)]
    for key, exp in expected.items():
        got = ctx.num(values, _summary_cell(ctx, key), LABELS[key])
        ok = _revenue_ok(values, got, exp, line_exacts) if key == "revenue" else X.close(got, exp)
        require(ok, f"{LABELS[key]} is {_amt(got)}, expected {_amt(exp)} (difference {_amt(got - exp, sign=True)})")


@rule(
    "summary-live",
    "Summary total revenue, total units and average unit price follow changes on Lines",
    kind="live",
    title_zh="Lines 表的数据变化时，Summary 的总收入、总数量和平均单价随之更新",
    requirement="...and keep updating when the Lines data changes.",
)
def summary_live(ctx):
    base, changed = ctx.values(), ctx.values(scenario="changed")
    rows = _changed(ctx.meta)
    expected = {
        "revenue": sum(price * qty * (1 - discount) for _, _, price, discount, qty in rows),
        "units": sum(qty for *_, qty in rows),
        "avg": sum(price for _, _, price, _, _ in rows) / len(rows),
    }
    line_exacts = [(r, price * qty * (1 - discount)) for r, _, price, discount, qty in rows]
    wrong = []
    for key, exp in expected.items():
        ref = _summary_cell(ctx, key)
        got = ctx.num(changed, ref, LABELS[key])
        ok = _revenue_ok(changed, got, exp, line_exacts) if key == "revenue" else X.close(got, exp)
        if ok:
            continue
        note = _same_as_before(base, ref, got)
        if key == "revenue":
            shown = [changed[X.make_ref(LINES, (r, 5))] for r, *_ in rows]
            if all(X.is_number(v) for v in shown) and X.close(sum(shown), got):
                note += ("; it does equal the sum of the Line Totals in that copy, so the gap comes from Line Totals "
                         "that are wrong in that copy (see line-totals-live)")
        wrong.append(f"{LABELS[key]} is {_amt(got)}, expected {_amt(exp)} (difference {_amt(got - exp, sign=True)}){note}")
    require(not wrong, f"{CHANGED}, " + "; ".join(wrong))


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
        got = vals.cell(r, 1).value
        require(got == line["product"], f"row {r} holds {X.show(got)}, expected {line['product']!r}")
        got = vals.cell(r, 2).value
        require(X.close(got, line["price"]), f"Unit Price of {line['product']} is {X.show(got)}, expected {line['price']}")
        got = vals.cell(r, 4).value
        require(X.close(got, line["qty"]), f"Qty of {line['product']} is {X.show(got)}, expected {line['qty']}")
