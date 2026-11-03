import openpyxl

from failbook import kit
from failbook import xlsx as X
from failbook.checks import Fail, Inconclusive, require, rule

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


def _new_rate(meta):
    return round(meta["rate"] + 0.05, 4)


def _pct(rate):
    return f"{rate * 100:g}%"


def _amt(x, sign=False):
    """A money amount with as many decimals as it needs (at least 2, at most 6)."""
    whole, _, frac = (f"{x:+,.6f}" if sign else f"{x:,.6f}").rstrip("0").partition(".")
    return f"{whole}.{frac.ljust(2, '0')}"


def _whole_cents(x):
    return abs(x * 100 - round(x * 100)) < 1e-6


def _money_ok(got, exact):
    """The exact amount, or that amount rounded to whole cents (in either direction).

    The task gives VAT = Net x rate and does not ask for unrounded values; rounding VAT
    to the cent, line by line, is normal on invoices, so it is accepted. "Within one cent"
    rather than "less than one cent" keeps a round-down of a product such as 305.00999...
    (the binary form of 305.01) from failing.
    """
    return X.close(got, exact) or (_whole_cents(got) and abs(got - exact) <= 0.01 + 1e-9)


def _read(ctx, values, problems, col, r, what):
    """A numeric cell, or None after recording why it cannot be used (if ``problems`` is given)."""
    try:
        return ctx.num(values, X.make_ref(SHEET, (r, col)), what)
    except (Fail, Inconclusive) as exc:
        if problems is not None:
            problems.append(exc)
    return None


def _check(ctx, rate, *, vat=False, gross=False, totals=False, scenario=None):
    """Compare the invoice with ``rate``. Returns (problems, number of values checked).

    A problem is a message (the value is wrong) or an Inconclusive (the checker cannot
    read it). In a scenario, a message also says when a value did not change at all.
    """
    values = ctx.values(scenario=scenario)
    base = ctx.values() if scenario else None

    def stale(col, r, got):
        if base is None:
            return ""
        before = base[X.make_ref(SHEET, (r, col))]
        return ", the same as before the change" if X.is_number(before) and X.close(before, got) else ""

    problems, line_vats, line_gross, checked = [], [], [], 0
    for r, line in zip(_rows(ctx), ctx.meta["lines"]):
        net, item = _net(line), line["item"]
        exp_vat = net * rate
        if vat or totals:
            checked += vat
            got = _read(ctx, values, problems if vat else None, 5, r, f"VAT of '{item}'")
            ok = got is not None and _money_ok(got, exp_vat)
            line_vats.append(got if ok else None)
            if vat and got is not None and not ok:
                problems.append(f"VAT of '{item}' (row {r}) is {_amt(got)}, expected {_amt(exp_vat)} "
                                f"(Net {_amt(net)} x {_pct(rate)}; difference {_amt(got - exp_vat, sign=True)})"
                                f"{stale(5, r, got)}")
        if gross or totals:
            checked += gross
            got = _read(ctx, values, problems if gross else None, 6, r, f"Gross of '{item}'")
            ok = got is not None and _money_ok(got - net, exp_vat)
            line_gross.append(got if ok else None)
            if gross and got is not None and not ok:
                exp = net + exp_vat
                problems.append(f"Gross of '{item}' (row {r}) is {_amt(got)}, expected {_amt(exp)} "
                                f"(Net {_amt(net)} + VAT {_amt(exp_vat)}; difference {_amt(got - exp, sign=True)})"
                                f"{stale(6, r, got)}")
    if not totals:
        return problems, checked
    checked += 3
    t = _total_row(ctx)
    net = sum(_net(line) for line in ctx.meta["lines"])
    exp_vat = net * rate
    got_net = _read(ctx, values, problems, 4, t, "Total Net")
    got_vat = _read(ctx, values, problems, 5, t, "Total VAT")
    got_gross = _read(ctx, values, problems, 6, t, "Total Gross")
    net_ok = got_net is not None and X.close(got_net, net)
    if got_net is not None and not net_ok:
        problems.append(f"Total Net is {_amt(got_net)}, expected {_amt(net)} "
                        f"(difference {_amt(got_net - net, sign=True)}){stale(4, t, got_net)}")
    # Totals consistent with lines rounded to the cent are accepted too.
    vat_sum = sum(line_vats) if None not in line_vats else None
    gross_sum = sum(line_gross) if None not in line_gross else None
    vat_ok = got_vat is not None and (_money_ok(got_vat, exp_vat) or (vat_sum is not None and X.close(got_vat, vat_sum)))
    if got_vat is not None and not vat_ok:
        also = f" or {_amt(vat_sum)} (the sum of the line VATs)" if vat_sum is not None else ""
        problems.append(f"Total VAT is {_amt(got_vat)}, expected {_amt(exp_vat)} (Net total {_amt(net)} x "
                        f"{_pct(rate)}){also}; difference {_amt(got_vat - exp_vat, sign=True)}{stale(5, t, got_vat)}")
    if got_gross is not None:
        exp = net + exp_vat
        ok = (_money_ok(got_gross - net, exp_vat) or (gross_sum is not None and X.close(got_gross, gross_sum))
              or (vat_ok and net_ok and X.close(got_gross, got_net + got_vat)))
        if not ok:
            also = f" or {_amt(gross_sum)} (the sum of the line Gross amounts)" if gross_sum is not None else ""
            problems.append(f"Total Gross is {_amt(got_gross)}, expected {_amt(exp)} (Net total + VAT total){also}; "
                            f"difference {_amt(got_gross - exp, sign=True)}{stale(6, t, got_gross)}")
    return problems, checked


def _report(problems, checked, lead=""):
    """Fail with the wrong values (rounding to whole cents is accepted), else surface an inconclusive read."""
    fails = [str(p) for p in problems if not isinstance(p, Inconclusive)]
    if fails:
        count = f"{len(fails)} of {checked} values are wrong, e.g. " if len(fails) > 1 else ""
        raise Fail(f"{lead}{count}{'; '.join(fails[:3])}. Exact amounts and amounts rounded to whole cents are "
                   "both accepted.")
    if problems:
        raise problems[0]


def scenarios(ctx):
    return {"rate_up": {X.make_ref(SHEET, "B1"): _new_rate(ctx.meta)}}


@rule(
    "vat-correct",
    "VAT = Net x rate on every line",
    kind="value",
    title_zh="每一行 VAT = Net × 税率",
    requirement="VAT = Net x the VAT rate in B1, for every line item.",
)
def vat_correct(ctx):
    _report(*_check(ctx, ctx.meta["rate"], vat=True))


@rule(
    "gross-correct",
    "Gross = Net + VAT on every line",
    kind="value",
    title_zh="每一行 Gross = Net + VAT",
    requirement="Gross = Net + VAT, for every line item.",
)
def gross_correct(ctx):
    _report(*_check(ctx, ctx.meta["rate"], gross=True))


@rule(
    "vat-follows-rate",
    "Changing the rate in B1 updates VAT and Gross on every line and the totals",
    kind="live",
    title_zh="修改 B1 的税率后，每一行的 VAT、Gross 以及合计都会更新",
    requirement="Use formulas, so that changing the rate in B1 updates every line and the totals.",
)
def vat_follows_rate(ctx):
    rate = _new_rate(ctx.meta)
    problems, checked = _check(ctx, rate, vat=True, gross=True, totals=True, scenario="rate_up")
    _report(problems, checked, lead=f"with B1 changed to {_pct(rate)} in a copy of the file, ")


@rule(
    "totals-correct",
    "The Total row shows the correct Net, VAT and Gross totals",
    kind="value",
    title_zh="Total 行的 Net、VAT、Gross 合计正确",
    requirement="The Total row at the bottom must show the correct Net, VAT and Gross totals.",
)
def totals_correct(ctx):
    _report(*_check(ctx, ctx.meta["rate"], totals=True))


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
