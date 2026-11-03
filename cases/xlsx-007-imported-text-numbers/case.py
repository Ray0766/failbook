import csv
import datetime as dt
import io

import openpyxl

from failbook import kit
from failbook import xlsx as X
from failbook.checks import require, rule

EXP, SUM = "Expenses", "Summary"
CATEGORIES = ["Travel", "Meals", "Office", "Software", "Training"]
MONTHS = [("January", 1), ("February", 2), ("March", 3)]
DESCRIPTIONS = {
    "Travel": ["Train ticket", "Taxi to airport", "Hotel, 2 nights", "Flight change fee"],
    "Meals": ["Client lunch", "Team dinner", "Coffee with partner"],
    "Office": ["Printer toner", "Desk chair", "Paper and pens"],
    "Software": ["Design tool licence", "Cloud storage", "Video calls plan"],
    "Training": ["Conference pass", "Online course", "Workshop fee"],
}


def params(seed):
    r = kit.rng(seed, "xlsx-007")
    existing = []
    for month in (1, 2):
        for _ in range(r.randint(6, 8)):
            cat = r.choice(CATEGORIES)
            existing.append({
                "date": dt.date(2024, month, r.randint(1, 28)).isoformat(),
                "category": cat,
                "description": r.choice(DESCRIPTIONS[cat]),
                "amount": round(r.uniform(12, 900), 2),
            })
    existing.sort(key=lambda e: e["date"])
    new = []
    days = sorted(r.sample(range(1, 29), 6))
    for i, day in enumerate(days):
        cat = CATEGORIES[i % len(CATEGORIES)] if i < 5 else r.choice(CATEGORIES)
        big = i in (0, 3)
        new.append({
            "date": dt.date(2024, 3, day).isoformat(),
            "category": cat,
            "description": r.choice(DESCRIPTIONS[cat]) + f" #{i + 1}",
            "amount": round(r.uniform(1000, 4800) if big else r.uniform(15, 600), 2),
        })
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\n")
    w.writerow(["date", "category", "description", "amount"])
    for e in new:
        w.writerow([e["date"], e["category"], e["description"], f"{e['amount']:,.2f}"])
    return {"existing": existing, "new": new, "csv": buf.getvalue().strip()}


def build(meta, outdir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = EXP
    ws.append(["Date", "Category", "Description", "Amount"])
    kit.style_header(ws, 1, 4)
    for i, e in enumerate(meta["existing"], start=2):
        ws.append([kit.from_iso(e["date"]), e["category"], e["description"], e["amount"]])
        ws.cell(i, 1).number_format = kit.DATE_FMT
        ws.cell(i, 4).number_format = kit.MONEY_FMT
    kit.set_widths(ws, {"A": 12, "B": 11, "C": 26, "D": 12})
    ws.freeze_panes = "A2"

    s = wb.create_sheet(SUM)
    s.append(["Category", "Total"])
    kit.style_header(s, 1, 2)
    for i, cat in enumerate(CATEGORIES, start=2):
        s.cell(i, 1, cat)
        s.cell(i, 2, f"=SUMIF(Expenses!B:B,A{i},Expenses!D:D)").number_format = kit.MONEY_FMT
    row = len(CATEGORIES) + 3
    s.cell(row, 1, "Month")
    s.cell(row, 2, "Total")
    kit.style_header(s, row, 2)
    for name, m in MONTHS:
        row += 1
        s.cell(row, 1, name)
        s.cell(row, 2, f'=SUMIFS(Expenses!D:D,Expenses!A:A,">="&DATE(2024,{m},1),Expenses!A:A,"<"&DATE(2024,{m + 1},1))')
        s.cell(row, 2).number_format = kit.MONEY_FMT
    row += 2
    s.cell(row, 1, "Grand total").font = kit.TOTAL_FONT
    s.cell(row, 2, "=SUM(Expenses!D:D)").number_format = kit.MONEY_FMT
    kit.set_widths(s, {"A": 14, "B": 14})
    wb.save(outdir / "expenses.xlsx")


def _new_rows(ctx):
    ws = ctx.ws(EXP)
    first_new = 2 + len(ctx.meta["existing"])
    rows = []
    for i, e in enumerate(ctx.meta["new"]):
        r = first_new + i
        got = ws.cell(r, 3).value
        require(X.norm_text(got) == X.norm_text(e["description"]),
                f"row {r} should hold the new expense '{e['description']}' but holds {X.show(got)}")
        rows.append((r, e))
    return rows


@rule(
    "rows-appended",
    "The new expenses are appended after the existing ones, in order",
    kind="content",
    title_zh="新费用按顺序追加在现有记录之后",
    requirement="Append these expenses from our bank export to the end of the Expenses sheet, in this order.",
)
def rows_appended(ctx):
    rows = _new_rows(ctx)
    ws = ctx.ws(EXP)
    for r, e in rows:
        require(X.norm_text(ws.cell(r, 2).value) == X.norm_text(e["category"]), f"row {r}: wrong category")
    after = rows[-1][0] + 1
    require(all(ws.cell(after, c).value in (None, "") for c in range(1, 5)), f"unexpected data in row {after}")


@rule(
    "amounts-are-numbers",
    "New amounts are stored as numbers with the right value",
    kind="content",
    title_zh="新金额存为数字，且数值正确",
    requirement="Store ... the amounts as real numbers (not text).",
)
def amounts_are_numbers(ctx):
    vals = ctx.values().ws(EXP)
    bad = []
    for r, e in _new_rows(ctx):
        got = vals.cell(r, 4).value
        if not X.is_number(got):
            bad.append(f"row {r}: {X.show(got)} is {'text' if isinstance(got, str) else type(got).__name__}")
        elif not X.close(got, e["amount"]):
            bad.append(f"row {r}: {got} instead of {e['amount']}")
    require(not bad, "amounts are wrong: " + "; ".join(bad[:3]))


@rule(
    "dates-are-dates",
    "New dates are stored as real dates with the right value",
    kind="content",
    title_zh="新日期存为真正的日期，且日期正确",
    requirement="Store the dates as real dates ... (not text).",
)
def dates_are_dates(ctx):
    vals = ctx.values().ws(EXP)
    bad = []
    for r, e in _new_rows(ctx):
        got = vals.cell(r, 1).value
        if not X.is_date(got):
            bad.append(f"row {r}: {X.show(got)} is not a date")
        elif X.as_date(got) != kit.from_iso(e["date"]):
            bad.append(f"row {r}: {X.show(got)} instead of {e['date']}")
    require(not bad, "dates are wrong: " + "; ".join(bad[:3]))


@rule(
    "summary-includes-new",
    "Every Summary total (category, month, grand total) includes the new expenses",
    kind="value",
    title_zh="Summary 的每项合计（类别、月份、总计）都包含新费用",
    requirement="...so that every total on the Summary sheet includes the new expenses.",
)
def summary_includes_new(ctx):
    values = ctx.values()
    ws = ctx.ws(SUM)
    every = ctx.meta["existing"] + ctx.meta["new"]
    expected = {cat: sum(e["amount"] for e in every if e["category"] == cat) for cat in CATEGORIES}
    for name, m in MONTHS:
        expected[name] = sum(e["amount"] for e in every if kit.from_iso(e["date"]).month == m)
    expected["Grand total"] = sum(e["amount"] for e in every)
    for label, exp in expected.items():
        hit = X.find_cell(ws, label, col=1)
        require(hit, f"Summary label {label!r} is missing")
        got = ctx.num(values, X.make_ref(SUM, (hit[0], 2)), f"Summary '{label}'")
        require(X.close(got, exp), f"Summary '{label}' is {got:,.2f}, expected {exp:,.2f} (missing {exp - got:,.2f})")


@rule(
    "existing-unchanged",
    "The existing expenses are unchanged",
    kind="preserve",
    title_zh="原有费用记录保持不变",
    requirement="Do not change anything the task does not ask you to change.",
)
def existing_unchanged(ctx):
    vals = ctx.values().ws(EXP)
    for r, e in enumerate(ctx.meta["existing"], start=2):
        got = [vals.cell(r, c).value for c in range(1, 5)]
        require(X.as_date(got[0]) == kit.from_iso(e["date"]) and got[1] == e["category"]
                and got[2] == e["description"] and X.close(got[3], e["amount"]), f"existing row {r} changed: {got}")
