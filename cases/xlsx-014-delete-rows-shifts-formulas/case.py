import openpyxl

from failbook import kit
from failbook import xlsx as X
from failbook.checks import require, rule

STOCK, SUMMARY = "Stock", "Summary"
HEADERS = ["SKU", "Item", "Status", "Qty", "Unit Cost", "Value"]
LABELS = {"active": "Active items", "units": "Total units", "value": "Total stock value"}


def params(seed):
    r = kit.rng(seed, "xlsx-014")
    n = r.randint(16, 22)
    names = kit.sample(r, kit.PRODUCTS, n)
    items = []
    for i, name in enumerate(names):
        items.append({
            "sku": f"S-{301 + i}",
            "item": name,
            "status": r.choices(["Active", "Backorder"], weights=[5, 1])[0],
            "qty": r.randint(0, 250),
            "cost": round(r.uniform(1.5, 140), 2),
        })
    # Discontinue a handful of items, never the first or the last, so later rows must move.
    for i in r.sample(range(1, n - 1), r.randint(4, 6)):
        items[i]["status"] = "Discontinued"
    return {"items": items}


def build(meta, outdir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = STOCK
    ws.append(HEADERS)
    kit.style_header(ws, 1, len(HEADERS))
    for r, it in enumerate(meta["items"], start=2):
        ws.append([it["sku"], it["item"], it["status"], it["qty"], it["cost"], f"=D{r}*E{r}"])
        ws.cell(r, 5).number_format = kit.MONEY_FMT
        ws.cell(r, 6).number_format = kit.MONEY_FMT
    last = 1 + len(meta["items"])
    total = last + 1
    ws.cell(total, 1, "Total").font = kit.TOTAL_FONT
    ws.cell(total, 4, f"=SUM(D2:D{last})").font = kit.TOTAL_FONT
    ws.cell(total, 6, f"=SUM(F2:F{last})").font = kit.TOTAL_FONT
    ws.cell(total, 6).number_format = kit.MONEY_FMT
    kit.set_widths(ws, {"A": 8, "B": 18, "C": 13, "D": 7, "E": 10, "F": 12})
    ws.freeze_panes = "A2"
    s = wb.create_sheet(SUMMARY)
    s["A1"], s["B1"] = LABELS["active"], f'=COUNTIF(Stock!C2:C{last},"Active")'
    s["A2"], s["B2"] = LABELS["units"], f"=Stock!D{total}"
    s["A3"], s["B3"] = LABELS["value"], f"=Stock!F{total}"
    s["B3"].number_format = kit.MONEY_FMT
    kit.set_widths(s, {"A": 20, "B": 14})
    wb.save(outdir / "inventory.xlsx")


def _kept(meta):
    return [it for it in meta["items"] if it["status"] != "Discontinued"]


def _kept_rows(ctx):
    return list(zip(range(2, 2 + len(_kept(ctx.meta))), _kept(ctx.meta)))


def _total_row(ctx):
    ws = ctx.ws(STOCK)
    hits = X.find_cells(ws, "Total", col=1)
    require(hits, "no row labelled 'Total' in the Stock sheet")
    return hits[-1][0]


def _summary_ref(ctx, key):
    hit = X.find_cell(ctx.ws(SUMMARY), LABELS[key], col=1)
    require(hit, f"label {LABELS[key]!r} not found on the Summary sheet")
    return X.make_ref(SUMMARY, (hit[0], 2))


def scenarios(ctx):
    rows = _kept_rows(ctx)
    return {
        "more": {X.make_ref(STOCK, (r, 4)): it["qty"] + 10 for r, it in rows},
        "dearer": {X.make_ref(STOCK, (r, 5)): round(it["cost"] + 1, 2) for r, it in rows},
    }


@rule(
    "discontinued-removed",
    "No Discontinued item is left, and no blank rows were left behind",
    kind="content",
    title_zh="不再有 Discontinued 商品，也没有留下空行",
    requirement="Remove every item whose Status is Discontinued. Delete the rows entirely; do not leave blank rows.",
)
def discontinued_removed(ctx):
    vals = ctx.values().ws(STOCK)
    total = _total_row(ctx)
    for r in range(2, total):
        status = vals.cell(r, 3).value
        require(X.norm_text(status) != "discontinued", f"row {r} ({vals.cell(r, 1).value}) is still Discontinued")
        require(vals.cell(r, 1).value not in (None, ""), f"row {r} is blank")
    require(total == 2 + len(_kept(ctx.meta)), f"the Total row is on row {total}, expected {2 + len(_kept(ctx.meta))}")


@rule(
    "remaining-items-kept",
    "All other items remain, in their original order, with unchanged data",
    kind="preserve",
    title_zh="其余商品全部保留，顺序和数据不变",
    requirement="Remove every item whose Status is Discontinued (and nothing else).",
)
def remaining_items_kept(ctx):
    vals = ctx.values().ws(STOCK)
    for r, it in _kept_rows(ctx):
        got = [vals.cell(r, c).value for c in range(1, 6)]
        require(got[:4] == [it["sku"], it["item"], it["status"], it["qty"]] and X.close(got[4], it["cost"]),
                f"row {r} is {got}, expected {it['sku']} {it['item']}")


@rule(
    "values-correct",
    "Value = Qty x Unit Cost on every remaining row",
    kind="value",
    title_zh="剩余每一行 Value = Qty × Unit Cost",
    requirement="The Value column ... must stay correct for the remaining items.",
)
def values_correct(ctx):
    values = ctx.values()
    bad = []
    for r, it in _kept_rows(ctx):
        got = ctx.num(values, X.make_ref(STOCK, (r, 6)), f"Value of {it['sku']}")
        if not X.close(got, it["qty"] * it["cost"]):
            bad.append(f"{it['sku']} (row {r}): {got:,.2f} instead of {it['qty'] * it['cost']:,.2f}")
    require(not bad, f"{len(bad)} Value cell(s) are wrong, e.g. " + "; ".join(bad[:3]))


@rule(
    "stock-totals-correct",
    "The Total row sums Qty and Value over the remaining items",
    kind="value",
    title_zh="Total 行对剩余商品的 Qty 与 Value 求和",
    requirement="The Total row ... must stay correct for the remaining items.",
)
def stock_totals_correct(ctx):
    values = ctx.values()
    total = _total_row(ctx)
    kept = _kept(ctx.meta)
    for col, label, exp in ((4, "Qty", sum(i["qty"] for i in kept)), (6, "Value", sum(i["qty"] * i["cost"] for i in kept))):
        got = ctx.num(values, X.make_ref(STOCK, (total, col)), f"Total {label}")
        require(X.close(got, exp), f"Total {label} is {got:,.2f}, expected {exp:,.2f}")


@rule(
    "summary-correct",
    "Summary shows the right active count, total units and total stock value",
    kind="value",
    title_zh="Summary 的在售商品数、总数量、库存总值正确",
    requirement="The Summary sheet must stay correct for the remaining items.",
)
def summary_correct(ctx):
    values = ctx.values()
    kept = _kept(ctx.meta)
    expected = {
        "active": sum(i["status"] == "Active" for i in kept),
        "units": sum(i["qty"] for i in kept),
        "value": sum(i["qty"] * i["cost"] for i in kept),
    }
    for key, exp in expected.items():
        got = ctx.num(values, _summary_ref(ctx, key), LABELS[key])
        require(X.close(got, exp), f"{LABELS[key]} is {got:,.2f}, expected {exp:,.2f}")


@rule(
    "totals-live",
    "Values, totals and Summary update when quantities or costs change",
    kind="live",
    title_zh="数量或成本变化时，Value、合计与 Summary 随之更新",
    requirement="...and keep updating (as formulas) when quantities or costs change.",
)
def totals_live(ctx):
    base = ctx.values()
    total = _total_row(ctx)
    for scenario, change, col in (("more", "adding 10 units to every remaining item", 4),
                                  ("dearer", "raising every remaining unit cost by 1", 5)):
        changed = ctx.values(scenario=scenario)
        d_qty = d_value = 0.0
        other_col = 5 if col == 4 else 4  # Value = Qty x Unit Cost
        for r, it in _kept_rows(ctx):
            ref_in, ref_other = X.make_ref(STOCK, (r, col)), X.make_ref(STOCK, (r, other_col))
            what = f"{HEADERS[col - 1]} of {it['sku']}"
            d_in = ctx.num(changed, ref_in, what) - ctx.num(base, ref_in, what)
            expected = d_in * ctx.num(base, ref_other, f"{HEADERS[other_col - 1]} of {it['sku']}")
            d_qty += d_in if col == 4 else 0.0
            d_value += expected
            ref = X.make_ref(STOCK, (r, 6))
            got = ctx.num(changed, ref, f"Value of {it['sku']}") - ctx.num(base, ref, f"Value of {it['sku']}")
            require(X.close(got, expected), f"{change} changed the Value of {it['sku']} (row {r}) by {got:+,.2f}, "
                    f"expected {expected:+,.2f}")
        for ref, label, exp in (
            (X.make_ref(STOCK, (total, 4)), "Total Qty", d_qty),
            (X.make_ref(STOCK, (total, 6)), "Total Value", d_value),
            (_summary_ref(ctx, "units"), LABELS["units"], d_qty),
            (_summary_ref(ctx, "value"), LABELS["value"], d_value),
        ):
            delta = ctx.num(changed, ref, label) - ctx.num(base, ref, label)
            require(X.close(delta, exp), f"{change} changed {label} by {delta:+,.2f}, expected {exp:+,.2f}")
