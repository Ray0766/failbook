import openpyxl

from failbook import kit
from failbook import xlsx as X
from failbook.checks import require, rule

SHEET = "Sales"
HEADERS = ["Region", "Rep", "Q1", "Q2", "Q3", "Q4", "Year"]
QCOLS = [3, 4, 5, 6]
YEAR = 7


def params(seed):
    r = kit.rng(seed, "xlsx-010")
    regions = kit.sample(r, kit.REGIONS, r.randint(3, 4))
    names = kit.people(r, 4 * len(regions))
    groups = []
    for region in regions:
        reps = []
        for _ in range(r.randint(2, 4)):
            reps.append({"rep": names.pop(), "q": [r.randrange(8000, 60000, 100) for _ in range(4)]})
        groups.append({"region": region, "reps": reps})
    return {"groups": groups}


def layout(meta):
    """Row numbers of each rep and subtotal in the input."""
    rows, subtotals, r = [], [], 2
    for g in meta["groups"]:
        start = r
        for rep in g["reps"]:
            rows.append((r, g["region"], rep))
            r += 1
        subtotals.append((r, g["region"], start, r - 1))
        r += 1
    return rows, subtotals


def build(meta, outdir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.append(HEADERS)
    kit.style_header(ws, 1, len(HEADERS))
    reps, subtotals = layout(meta)
    for r, region, rep in reps:
        ws.cell(r, 1, region)
        ws.cell(r, 2, rep["rep"])
        for c, v in zip(QCOLS, rep["q"]):
            ws.cell(r, c, v).number_format = kit.INT_FMT
        ws.cell(r, YEAR, f"=SUM(C{r}:F{r})").number_format = kit.INT_FMT
    for r, region, start, end in subtotals:
        ws.cell(r, 1, f"{region} subtotal").font = kit.TOTAL_FONT
        for col in "CDEFG":
            cell = ws[f"{col}{r}"]
            cell.value = f"=SUM({col}{start}:{col}{end})"
            cell.font = kit.TOTAL_FONT
            cell.number_format = kit.INT_FMT
    kit.set_widths(ws, {"A": 18, "B": 16, "C": 10, "D": 10, "E": 10, "F": 10, "G": 11})
    ws.freeze_panes = "A2"
    wb.save(outdir / "regional.xlsx")


def _grand_row(ctx):
    ws = ctx.ws(SHEET)
    hits = X.find_cells(ws, "Grand Total", col=1)
    require(hits, "no row labelled 'Grand Total' in column A")
    require(len(hits) == 1, f"{len(hits)} rows are labelled 'Grand Total'")
    return hits[0][0]


def _expected(ctx, bump=0):
    reps, _ = layout(ctx.meta)
    q = [sum(rep["q"][i] for _, _, rep in reps) for i in range(4)]
    q[1] += bump
    return q + [sum(q)]


def _bumped(meta):
    """The "bump" scenario: (row, column, new value) adding 100 to one rep's figure in each quarter.

    Each quarter uses a different rep, so a Grand Total that covers only some rows is caught too.
    """
    reps, _ = layout(meta)
    out = []
    for i, col in enumerate(QCOLS):
        r, _, rep = reps[(len(reps) // 2 + i) % len(reps)]
        out.append((r, col, rep["q"][i] + 100))
    return out


def scenarios(ctx):
    return {"bump": {X.make_ref(SHEET, (r, col)): value for r, col, value in _bumped(ctx.meta)}}


@rule(
    "grand-total-row",
    "A 'Grand Total' row sits directly below the last subtotal",
    kind="content",
    title_zh="'Grand Total' 行位于最后一个小计行的正下方",
    requirement="Add a Grand Total row directly below the last subtotal row, with the label 'Grand Total' in column A.",
)
def grand_total_row(ctx):
    _, subtotals = layout(ctx.meta)
    expected = subtotals[-1][0] + 1
    got = _grand_row(ctx)
    require(got == expected, f"the Grand Total is on row {got}, expected row {expected}")


@rule(
    "grand-total-values",
    "Grand Total equals the sum over all reps (subtotals not counted twice)",
    kind="value",
    title_zh="Grand Total 等于全部销售代表的合计（小计未被重复计入）",
    requirement="For Q1, Q2, Q3, Q4 and Year, it must show the total over all reps in all regions.",
)
def grand_total_values(ctx):
    values = ctx.values()
    row = _grand_row(ctx)
    for col, exp, label in zip([*QCOLS, YEAR], _expected(ctx), HEADERS[2:]):
        got = ctx.num(values, X.make_ref(SHEET, (row, col)), f"Grand Total {label}")
        hint = " (exactly double: the subtotals are counted too)" if X.close(got, 2 * exp) else ""
        require(X.close(got, exp), f"Grand Total {label} is {got:,.0f}, expected {exp:,.0f}{hint}")


@rule(
    "grand-total-live",
    "Adding 100 to one rep's figure in each quarter adds exactly 100 to each quarter's Grand Total and 400 to Year",
    kind="live",
    title_zh="每个季度各给一位代表加 100，各季度 Grand Total 恰好增加 100，Year 增加 400",
    requirement="For Q1, Q2, Q3, Q4 and Year ... Use formulas, so that it updates when a rep's figures change.",
)
def grand_total_live(ctx):
    base, bumped = ctx.values(), ctx.values(scenario="bump")
    row = _grand_row(ctx)
    changed = {col: r for r, col, _ in _bumped(ctx.meta)}
    for col, label in [(c, HEADERS[c - 1]) for c in QCOLS] + [(YEAR, "Year")]:
        ref = X.make_ref(SHEET, (row, col))
        delta = ctx.num(bumped, ref, f"Grand Total {label}") - ctx.num(base, ref, f"Grand Total {label}")
        if col == YEAR:
            require(X.close(delta, 400), "adding 100 to one rep's figure in each of Q1-Q4 changed Grand Total Year "
                    f"by {delta:,.0f}, expected +400")
        else:
            require(X.close(delta, 100), f"adding 100 to one rep's {label} (row {changed[col]}) changed Grand Total "
                    f"{label} by {delta:,.0f}, expected +100")


@rule(
    "sales-unchanged",
    "Rep rows and subtotal rows are unchanged",
    kind="preserve",
    title_zh="销售代表行与小计行保持不变",
    requirement="Do not change anything the task does not ask you to change.",
)
def sales_unchanged(ctx):
    values = ctx.values()
    vals = values.ws(SHEET)
    reps, subtotals = layout(ctx.meta)
    for r, region, rep in reps:
        got = [vals.cell(r, c).value for c in (1, 2, *QCOLS)]
        require(got == [region, rep["rep"], *rep["q"]], f"row {r} changed: {got}")
    for r, region, start, end in subtotals:
        require(X.norm_text(vals.cell(r, 1).value) == X.norm_text(f"{region} subtotal"), f"subtotal row {r} moved")
        for i, c in enumerate(QCOLS):
            exp = sum(rep["q"][i] for rr, _, rep in reps if start <= rr <= end)
            got = ctx.num(values, X.make_ref(SHEET, (r, c)), f"{region} subtotal")
            require(X.close(got, exp), f"{region} subtotal {HEADERS[c - 1]} is {got:,.0f}, expected {exp:,.0f}")
