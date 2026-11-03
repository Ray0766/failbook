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


#: The "lines_changed" scenario raises, in a copy of the file, the k-th line's Qty by k (k = 1, 2,
#: ...) and every Unit Price by PRICE_BUMP, so a Net or a Total Net that is a typed number no
#: longer matches.
PRICE_BUMP = 1.0


def _changed_lines(meta):
    return [dict(line, qty=line["qty"] + k, price=round(line["price"] + PRICE_BUMP, 2))
            for k, line in enumerate(meta["lines"], start=1)]


def scenarios(ctx):
    lines = {}
    for r, line in zip(_rows(ctx), _changed_lines(ctx.meta)):
        lines[X.make_ref(SHEET, (r, 2))] = line["qty"]
        lines[X.make_ref(SHEET, (r, 3))] = line["price"]
    return {"rate_up": {X.make_ref(SHEET, "B1"): _new_rate(ctx.meta)}, "lines_changed": lines}


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


#: How the "lines_changed" copy is described in messages.
_LINES_CHANGED = (f"in a copy of the file where every line's Qty was raised by a different amount and every Unit "
                  f"Price by {PRICE_BUMP:.2f}")


@rule(
    "vat-gross-follow-net",
    "When Qty and Unit Price change, VAT and Gross on every line and in the Total row follow the Nets",
    kind="live",
    title_zh="数量和单价变化时，每一行以及 Total 行的 VAT 和 Gross 都随 Net 更新",
    requirement="Fill in, for every line item: VAT = Net x the VAT rate in B1; Gross = Net + VAT. Use formulas. "
                "The Total row at the bottom must show the correct Net, VAT and Gross totals.",
)
def vat_gross_follow_net(ctx):
    # Each figure is compared with the Nets as they are in the changed copy, so a Net that did not
    # follow its Qty and Unit Price (net-still-calculated) is not counted again here. A Net that is
    # not a number there is not reported either: that line's VAT and Gross are not compared, and
    # neither is a Total VAT or Total Gross that is not a number (it may carry the error along).
    base, changed = ctx.values(), ctx.values(scenario="lines_changed")
    rate = ctx.meta["rate"]
    problems, vats, grosses, bad_nets, bad_rows = [], [], [], [], []

    def stale(col, r, got):
        before = base[X.make_ref(SHEET, (r, col))]
        return ", the same as before the change" if X.is_number(before) and X.close(before, got) else ""

    def expected(options):
        """'a (how) or b (how)' for the accepted (value, accepted?, how) options, equal values merged."""
        merged = []
        for v, _, how in options:
            if v is None:
                continue
            same = next((m for m in merged if X.close(m[0], v)), None)
            if same:
                same[1].append(how)
            else:
                merged.append((v, [how]))
        return " or ".join(f"{_amt(v)} ({', and '.join(hows)})" for v, hows in merged)

    for r, line in zip(_rows(ctx), _changed_lines(ctx.meta)):
        item = line["item"]
        net = _read(ctx, changed, None, 4, r, f"Net of '{item}'")
        if net is None:  # its VAT and Gross may carry it (#VALUE!): not compared
            bad_nets.append(f"Net of '{item}' (row {r})")
            bad_rows.append((r, item))
            continue
        vat = _read(ctx, changed, problems, 5, r, f"VAT of '{item}'")
        gross = _read(ctx, changed, problems, 6, r, f"Gross of '{item}'")
        vats.append(vat)
        grosses.append(gross)
        exp_vat = net * rate
        if vat is not None and not _money_ok(vat, exp_vat):
            problems.append(f"VAT of '{item}' (row {r}) is {_amt(vat)}, expected {_amt(exp_vat)} (its Net in that "
                            f"copy {_amt(net)} x {_pct(rate)}){stale(5, r, vat)}")
        follows = _money_ok(gross - net, exp_vat) if gross is not None else True
        if not follows and not (vat is not None and X.close(gross, net + vat)):  # Gross = Net + VAT as shown
            problems.append(f"Gross of '{item}' (row {r}) is {_amt(gross)}, expected {_amt(net + exp_vat)} (its Net "
                            f"in that copy {_amt(net)} + VAT {_amt(exp_vat)}){stale(6, r, gross)}")
    # When every line's Net is a number, a total that is not a number is reported (unless Total Net is
    # not a number either). When some line's Net is not a number, a total that is not a number may
    # carry that error along and is not compared; a total that is a number still is.
    t, lines_read = _total_row(ctx), not bad_nets
    net = _read(ctx, changed, None, 4, t, "Total Net")
    if lines_read:
        vat = _read(ctx, changed, problems if net is not None else None, 5, t, "Total VAT")
        gross = _read(ctx, changed, problems if net is not None else None, 6, t, "Total Gross")
        if net is None and None in (vat, gross):
            problems.append(Inconclusive(f"{_LINES_CHANGED}, Total Net (row {t}) is not a number, so a Total VAT "
                                         "or Total Gross that is not a number either cannot be compared"))
        # A total is accepted if it is the sum of its column's lines as read in that copy, or follows Total Net.
        vat_sum = sum(vats) if None not in vats else None
        gross_sum = sum(grosses) if None not in grosses else None
        sum_how = "the sum of the line {} in that copy"
    else:
        vat = _read(ctx, changed, None, 5, t, "Total VAT")
        gross = _read(ctx, changed, None, 6, t, "Total Gross")
        for r, item in bad_rows:
            vats.append(_read(ctx, changed, None, 5, r, f"VAT of '{item}'"))
            grosses.append(_read(ctx, changed, None, 6, r, f"Gross of '{item}'"))
        vat_sum = sum(v for v in vats if v is not None)
        gross_sum = sum(g for g in grosses if g is not None)
        sum_how = "the sum of the line {} that are numbers in that copy"
    # A total worked out from Qty and Unit Price directly (SUMPRODUCT) follows them too, even where a
    # line's Net did not; that Net is reported under net-still-calculated, not here.
    lines_net = sum(_net(line) for line in _changed_lines(ctx.meta))
    if vat is not None:
        ok = [(vat_sum, X.close(vat, vat_sum) if vat_sum is not None else False, sum_how.format("VATs"))]
        if net is not None:
            ok.append((net * rate, _money_ok(vat, net * rate), f"Total Net in that copy x {_pct(rate)}"))
        ok.append((lines_net * rate, _money_ok(vat, lines_net * rate),
                   f"the sum of Qty x Unit Price in that copy x {_pct(rate)}"))
        if not any(good for _, good, _ in ok):
            problems.append(f"Total VAT (row {t}) is {_amt(vat)}, expected {expected(ok)}{stale(5, t, vat)}")
    if gross is not None:
        ok = [(gross_sum, X.close(gross, gross_sum) if gross_sum is not None else False,
               sum_how.format("Gross amounts"))]
        ok.append((lines_net * (1 + rate), _money_ok(gross - lines_net, lines_net * rate),
                   f"the sum of Qty x Unit Price in that copy + {_pct(rate)} of it"))
        if net is not None:
            ok.append((net + net * rate, _money_ok(gross - net, net * rate),
                       f"Total Net in that copy + {_pct(rate)} of it"))
            if vat is not None:
                ok.append((net + vat, X.close(gross, net + vat), "Total Net + Total VAT in that copy"))
        if not any(good for _, good, _ in ok):
            problems.append(f"Total Gross (row {t}) is {_amt(gross)}, expected {expected(ok)}{stale(6, t, gross)}")
    fails = [str(p) for p in problems if not isinstance(p, Inconclusive)]
    unread = [name for name, v in (("Total VAT", vat), ("Total Gross", gross)) if v is None]
    if lines_read and net is not None:
        checked, scope = 2 * len(vats) + 2, "the lines and the Total row"
    elif lines_read:  # Total Net is not a number: a total that is not a number either was not compared
        checked = 2 * len(vats) + 2 - len(unread)
        scope = "the lines" + (", and the Total row's figures that are numbers" if len(unread) < 2 else "")
    else:
        checked = 2 * (len(vats) - len(bad_rows)) + 2 - len(unread)
        scope = ("the lines whose Net in that copy is a number"
                 + (", and the Total row's figures that are numbers" if len(unread) < 2 else ""))
    require(not fails, f"{_LINES_CHANGED} (rate unchanged), {len(fails)} of {checked} VAT and Gross figures "
                       f"({scope}) did not follow the Nets: " + "; ".join(fails[:3])
                       + (f" (and {len(fails) - 3} more)" if len(fails) > 3 else "")
                       + ". Exact amounts and amounts rounded to whole cents are both accepted.")
    if problems:
        raise problems[0]
    if bad_nets:
        one = len(bad_nets) == 1
        is_, lines, nets = ("is", "that line's", "that Net") if one else ("are", "those lines'", "those Nets")
        totals = (f", nor {' and '.join(unread)}, which {'is not a number' if len(unread) == 1 else 'are not numbers'} "
                  "there" if unread else "")
        raise Inconclusive(f"{_LINES_CHANGED}, " + ", ".join(bad_nets[:3])
                           + (f" and {len(bad_nets) - 3} more" if len(bad_nets) > 3 else "")
                           + f" {is_} not a number, so {lines} VAT and Gross were not compared{totals} "
                           f"(net-still-calculated reports {nets})")


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
    "net-still-calculated",
    "Each line's Net still follows its Qty and Unit Price, and Total Net still follows the lines",
    kind="preserve",
    title_zh="每一行的 Net 仍随数量和单价变化，Net 合计仍随各行变化",
    requirement="Do not change anything the task does not ask you to change (in the file as given, each Net is "
                "calculated from its Qty and Unit Price, and Total Net from the lines).",
)
def net_still_calculated(ctx):
    base, changed = ctx.values(), ctx.values(scenario="lines_changed")
    lines = _changed_lines(ctx.meta)
    problems = []

    def problem(r, what, got, exp, how):
        before = base[X.make_ref(SHEET, (r, 4))]
        same = ", the same as before the change" if X.is_number(before) and X.close(before, got) else ""
        problems.append(f"{what} (row {r}) is {_amt(got)}, expected {_amt(exp)} ({how}){same}")

    nets = []
    for r, line in zip(_rows(ctx), lines):
        what = f"Net of '{line['item']}'"
        got = _read(ctx, changed, problems, 4, r, what)
        nets.append(got)
        if got is not None and not X.close(got, _net(line)):
            problem(r, what, got, _net(line), f"Qty {line['qty']} x Unit Price {_amt(line['price'])}")
    t, note = _total_row(ctx), ""
    got, exp = _read(ctx, changed, problems, 4, t, "Total Net"), sum(_net(line) for line in lines)
    if got is not None and not X.close(got, exp):
        if None not in nets and X.close(got, sum(nets)):
            # Total Net did follow the lines; the gap is in line Nets reported above.
            note = (f"; Total Net (row {t}) does equal the sum of the line Nets in that copy, "
                    f"{_amt(got)}, so it differs from {_amt(exp)} (the sum of Qty x Unit Price over the lines) "
                    "only because of those Nets")
        else:
            problem(t, "Total Net", got, exp, "the sum of Qty x Unit Price over the lines")
    fails = [str(p) for p in problems if not isinstance(p, Inconclusive)]
    checked = len(lines) + 1
    require(not fails, f"{_LINES_CHANGED}, {len(fails)} of {checked} Net figures (the lines and Total Net) "
                       "did not follow: "
                       + "; ".join(fails[:3]) + (f" (and {len(fails) - 3} more)" if len(fails) > 3 else "") + note)
    if problems:
        raise problems[0]


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
