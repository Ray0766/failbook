import openpyxl

from failbook import kit
from failbook import xlsx as X
from failbook.checks import require, rule

DATA, SUMMARY = "Q3 Data", "Summary"
MONTHS = ["Jul", "Aug", "Sep"]


def params(seed):
    r = kit.rng(seed, "xlsx-015")
    depts = kit.sample(r, kit.DEPARTMENTS, 5)
    rows = [{"dept": d, "month": m, "spend": r.randrange(5_000, 80_000, 10)} for d in depts for m in MONTHS]
    corrections = []
    for row in r.sample(rows, 3):
        new = row["spend"] + r.choice([-1, 1]) * r.randrange(1_000, 9_000, 10)
        corrections.append({"dept": row["dept"], "month": row["month"], "old": row["spend"], "new": new})
    table = kit.md_table(["Department", "Month", "Wrong value", "Correct value"],
                         [(c["dept"], c["month"], f"{c['old']:,}", f"{c['new']:,}") for c in corrections])
    return {"depts": depts, "rows": rows, "corrections": corrections, "corrections_md": table}


def _corrected(meta):
    fixes = {(c["dept"], c["month"]): c["new"] for c in meta["corrections"]}
    return [dict(row, spend=fixes.get((row["dept"], row["month"]), row["spend"])) for row in meta["rows"]]


def build(meta, outdir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = DATA
    ws.append(["Department", "Month", "Spend"])
    kit.style_header(ws, 1, 3)
    for i, row in enumerate(meta["rows"], start=2):
        ws.append([row["dept"], row["month"], row["spend"]])
        ws.cell(i, 3).number_format = kit.INT_FMT
    kit.set_widths(ws, {"A": 14, "B": 8, "C": 12})
    s = wb.create_sheet(SUMMARY)
    s.append(["Department", "Q3 total"])
    kit.style_header(s, 1, 2)
    # The totals were pasted in as values at some point: no formulas.
    for i, dept in enumerate(meta["depts"], start=2):
        s.cell(i, 1, dept)
        s.cell(i, 2, sum(row["spend"] for row in meta["rows"] if row["dept"] == dept)).number_format = kit.INT_FMT
    t = 2 + len(meta["depts"])
    s.cell(t, 1, "Total").font = kit.TOTAL_FONT
    s.cell(t, 2, sum(row["spend"] for row in meta["rows"])).number_format = kit.INT_FMT
    s.cell(t, 2).font = kit.TOTAL_FONT
    kit.set_widths(s, {"A": 14, "B": 12})
    wb.save(outdir / "budget.xlsx")


@rule(
    "corrections-applied",
    "The three corrections are applied in Q3 Data",
    kind="content",
    title_zh="Q3 Data 中的三处更正已完成",
    requirement="Apply these corrections in Q3 Data.",
)
def corrections_applied(ctx):
    vals = ctx.values().ws(DATA)
    for c in ctx.meta["corrections"]:
        hits = [r for r in range(2, vals.max_row + 1)
                if (vals.cell(r, 1).value, vals.cell(r, 2).value) == (c["dept"], c["month"])]
        require(hits, f"there is no {c['dept']} {c['month']} row in {DATA} any more")
        require(len(hits) == 1, f"{c['dept']} {c['month']} appears on {len(hits)} rows ({hits})")
        got = vals.cell(hits[0], 3).value
        require(X.close(got, c["new"]), f"{c['dept']} {c['month']} (row {hits[0]}) is {X.show(got)}, "
                f"expected the corrected {c['new']:,}")


@rule(
    "summary-dept-totals",
    "Each department's Q3 total on Summary reflects the corrections",
    kind="value",
    title_zh="Summary 上各部门的 Q3 合计已反映更正",
    requirement="Make sure the Summary sheet shows the corrected totals ... per department.",
)
def summary_dept_totals(ctx):
    values = ctx.values()
    ws = ctx.ws(SUMMARY)
    rows = _corrected(ctx.meta)
    for dept in ctx.meta["depts"]:
        hit = X.find_cell(ws, dept, col=1)
        require(hit, f"department {dept!r} is missing from the Summary")
        got = ctx.num(values, X.make_ref(SUMMARY, (hit[0], 2)), f"Summary total of {dept}")
        exp = sum(r["spend"] for r in rows if r["dept"] == dept)
        require(X.close(got, exp), f"Summary total of {dept} is {got:,.0f}, expected {exp:,.0f}")


@rule(
    "summary-grand-total",
    "The Summary grand total reflects the corrections",
    kind="value",
    title_zh="Summary 的总计已反映更正",
    requirement="Make sure the Summary sheet shows the corrected ... grand total.",
)
def summary_grand_total(ctx):
    values = ctx.values()
    hits = X.find_cells(ctx.ws(SUMMARY), "Total", col=1)
    require(hits, "the Total row is missing from the Summary")
    got = ctx.num(values, X.make_ref(SUMMARY, (hits[-1][0], 2)), "Summary grand total")
    exp = sum(r["spend"] for r in _corrected(ctx.meta))
    require(X.close(got, exp), f"Summary grand total is {got:,.0f}, expected {exp:,.0f}")


@rule(
    "summary-layout-kept",
    "Summary keeps its departments in the same order, followed by Total",
    kind="preserve",
    title_zh="Summary 部门顺序不变，末尾为 Total",
    requirement="Keep the Summary layout (department order and labels) as it is.",
)
def summary_layout_kept(ctx):
    ws = ctx.ws(SUMMARY)
    got = [ws.cell(r, 1).value for r in range(1, 3 + len(ctx.meta["depts"]))]
    exp = ["Department", *ctx.meta["depts"], "Total"]
    require([X.norm_text(v) for v in got] == [X.norm_text(v) for v in exp], f"Summary column A is {got}, expected {exp}")


@rule(
    "other-data-unchanged",
    "Rows that were not corrected are unchanged",
    kind="preserve",
    title_zh="未被更正的行保持不变",
    requirement="Do not change anything the task does not ask you to change.",
)
def other_data_unchanged(ctx):
    vals = ctx.values().ws(DATA)
    fixed = {(c["dept"], c["month"]) for c in ctx.meta["corrections"]}
    for i, row in enumerate(ctx.meta["rows"], start=2):
        got = (vals.cell(i, 1).value, vals.cell(i, 2).value, vals.cell(i, 3).value)
        require(got[:2] == (row["dept"], row["month"]),
                f"row {i} is {got[:2]}, expected {row['dept']} {row['month']} (rows must stay in place)")
        if (row["dept"], row["month"]) not in fixed:
            require(X.close(got[2], row["spend"]), f"{row['dept']} {row['month']} (row {i}) is {X.show(got[2])}, "
                    f"expected {row['spend']:,}; it is not one of the corrections")
    require(vals.cell(2 + len(ctx.meta["rows"]), 1).value in (None, ""), "there are extra rows below the data")
