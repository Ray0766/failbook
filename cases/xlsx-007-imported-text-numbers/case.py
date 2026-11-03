import csv
import datetime as dt
import io

import openpyxl

from failbook import kit
from failbook import xlsx as X
from failbook.checks import Inconclusive, require, rule

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


def _amt(x, places=2, sign=False):
    """A number with at least ``places`` and at most 6 decimals, so that two numbers the checker
    tells apart never print the same."""
    whole, _, frac = (f"{x:+,.6f}" if sign else f"{x:,.6f}").rstrip("0").partition(".")
    frac = frac.ljust(places, "0")
    return f"{whole}.{frac}" if frac else whole


def _reshuffled(meta):
    """[(row, original expense, new date, new category, new amount)] in the "reshuffled" scenario.

    In a copy of the file, every existing expense gets a different amount (20,000 or more), a
    category other than its own and a month other than its own (same day), so every total a row
    leaves or joins moves. The existing expenses (at least six in January and six in February)
    alternate, per month, between the two other months, so Jan, Feb and Mar all get rows; each row
    takes, in turn, a category that is not its own, preferring one no row has yet, so every
    category gets rows too (checked for seeds 0-1999). The new rows the agent appended are left as
    they are, so the change of each total does not depend on them.
    """
    out, not_given, per_month = [], list(CATEGORIES), {}
    for i, e in enumerate(meta["existing"]):
        old = kit.from_iso(e["date"])
        months = [m for _, m in MONTHS if m != old.month]
        k = per_month[old.month] = per_month.get(old.month, -1) + 1
        turn = CATEGORIES[i % len(CATEGORIES):] + CATEGORIES[:i % len(CATEGORIES)]
        choices = [c for c in turn if c != e["category"]]
        category = next((c for c in choices if c in not_given), choices[0])
        if category in not_given:
            not_given.remove(category)
        out.append((2 + i, e, dt.date(2024, months[k % len(months)], old.day), category, 20000 + 37.25 * (i + 1)))
    return out


def scenarios(ctx):
    overrides = {}
    for r, _, date, category, amount in _reshuffled(ctx.meta):
        overrides[X.make_ref(EXP, (r, 1))] = date
        overrides[X.make_ref(EXP, (r, 2))] = category
        overrides[X.make_ref(EXP, (r, 4))] = amount
    return {"reshuffled": overrides}


def _as_rows(expenses):
    return [(kit.from_iso(e["date"]), e["category"], e["amount"]) for e in expenses]


def _totals(expenses):
    """Summary label -> total, for a list of (date, category, amount)."""
    out = {cat: sum(a for _, c, a in expenses if c == cat) for cat in CATEGORIES}
    for name, m in MONTHS:
        out[name] = sum(a for d, _, a in expenses if d.month == m)
    out["Grand total"] = sum(a for _, _, a in expenses)
    return out


def _summary_refs(ctx):
    ws = ctx.ws(SUM)
    refs = {}
    for label in [*CATEGORIES, *(name for name, _ in MONTHS), "Grand total"]:
        hit = X.find_cell(ws, label, col=1)
        require(hit, f"Summary label {label!r} is missing")
        refs[label] = X.make_ref(SUM, (hit[0], 2))
    return refs


#: What X.number_format_key counts as one format, as the format messages say it.
_SAME_FORMAT = ("(formats count as the same when they differ only in letter case, quotes or backslash escapes, a "
                "text section that is only '@', a zero section equal to the first, a negative section of '-' and a "
                "first section of digits such as #,##0.00;-#,##0.00 (it shows -0.00 instead of 0.00 only for a "
                "negative number that rounds to zero, and the task's amounts are all positive), or a locale tag "
                "such as [$-409] on a date of digits; under such a tag an unquoted '/' counts as the date "
                "separator LibreOffice 24.2 shows for it: [$-41D]yyyy/mm/dd shows 2024-03-02, [$-409]yyyy/mm/dd "
                "2024/03/02 and [$-407]yyyy/mm/dd 2024.03.02).")


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
        got = ws.cell(r, 2).value
        require(X.norm_text(got) == X.norm_text(e["category"]),
                f"row {r} ({e['description']}): category is {X.show(got)}, expected {e['category']!r}")
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
    "new-rows-formatted",
    "The new dates and amounts have the number formats of the existing rows",
    kind="structure",
    title_zh="新记录的日期和金额与现有行的数字格式相同",
    requirement="Store the dates as real dates and the amounts as real numbers (not text), formatted like the "
                "existing rows.",
)
def new_rows_formatted(ctx):
    rows = _new_rows(ctx)
    ws, given = ctx.ws(EXP), ctx.input_wb()[EXP]
    # In the file as given the existing expenses are rows 2 onwards, all formatted alike. Columns
    # they leave in General (Category, Description) are not compared.
    general = X.number_format_key("General")
    wanted = {c: X.number_format(given, 2, c) for c in range(1, 5)}
    wanted = {c: f for c, f in wanted.items() if X.number_format_key(f) != general}
    names = {c: given.cell(1, c).value for c in wanted}
    wrong, unsure, why = [], [], {}  # why: column -> why the checker cannot compare it
    for r, e in rows:
        diff, maybe = [], []
        for c, f in wanted.items():
            got = X.number_format(ws, r, c)
            if X.number_format_key(got) != X.number_format_key(f):
                reason = X.number_format_unsure(got, f)  # '/' under a locale tag: the separator is unknown
                (maybe if reason else diff).append(f"{names[c]} '{got}'")
                if reason:
                    why[c] = reason
        if diff:
            wrong.append(f"row {r} ({e['description']}) has " + ", ".join(diff))
        elif maybe:
            unsure.append(f"row {r} ({e['description']}) has " + ", ".join(maybe))
    given_text = ", ".join(f"{names[c]} '{f}'" for c, f in wanted.items())
    require(not wrong, f"{len(wrong)} of {len(rows)} new rows are not formatted like the existing rows: "
                       + "; ".join(wrong[:3]) + (f" (and {len(wrong) - 3} more)" if len(wrong) > 3 else "")
                       + f". In {ctx.case.inputs[0]} as given, the existing rows have {given_text} {_SAME_FORMAT}")
    if unsure:
        reason = next(iter(why.values()))
        raise Inconclusive(f"{len(unsure)} of {len(rows)} new rows have a format that matches the existing rows' "
                           "except for '/' under a locale tag: " + "; ".join(unsure[:3])
                           + (f" (and {len(unsure) - 3} more)" if len(unsure) > 3 else "")
                           + f". {reason[0].upper()}{reason[1:]}, so it cannot tell whether these cells show like "
                           + ", ".join(f"{names[c]} '{wanted[c]}'" for c in why) + " in the existing rows")


@rule(
    "summary-includes-new",
    "Every Summary total (category, month, grand total) includes the new expenses",
    kind="value",
    title_zh="Summary 的每项合计（类别、月份、总计）都包含新费用",
    requirement="...so that every total on the Summary sheet includes the new expenses.",
)
def summary_includes_new(ctx):
    values = ctx.values()
    expected = _totals(_as_rows(ctx.meta["existing"] + ctx.meta["new"]))
    new_only = _totals(_as_rows(ctx.meta["new"]))
    refs = _summary_refs(ctx)
    wrong = []
    for label, ref in refs.items():
        got = ctx.num(values, ref, f"Summary '{label}'")
        exp = expected[label]
        if X.close(got, exp):
            continue
        note = ""
        if new_only[label] and X.close(exp - got, new_only[label], abs_tol=0.005):
            note = ", short by exactly the new expenses it should include"
        wrong.append(f"'{label}' is {_amt(got)}, expected {_amt(exp)} (difference {_amt(got - exp, sign=True)}){note}")
    require(not wrong, f"{len(wrong)} of {len(refs)} Summary totals are wrong: " + "; ".join(wrong[:3])
            + (f" (and {len(wrong) - 3} more)" if len(wrong) > 3 else ""))


@rule(
    "summary-still-calculated",
    "Every Summary total still follows changes of the expenses (category, month, grand total)",
    kind="preserve",
    title_zh="Summary 的每项合计（类别、月份、总计）仍随费用数据变化",
    requirement="Do not change anything the task does not ask you to change (the Summary sheet totals the "
                "Expenses sheet by category, by month and overall).",
)
def summary_still_calculated(ctx):
    base, changed = ctx.values(), ctx.values(scenario="reshuffled")
    rows = _reshuffled(ctx.meta)
    before = _totals(_as_rows(e for _, e, *_ in rows))
    after = _totals([(date, category, amount) for _, _, date, category, amount in rows])
    refs = _summary_refs(ctx)
    wrong = []
    for label, ref in refs.items():
        moved = ctx.num(changed, ref, f"Summary '{label}'") - ctx.num(base, ref, f"Summary '{label}'")
        exp = after[label] - before[label]
        if not X.close(moved, exp, abs_tol=0.005):
            same = ", so it did not change at all" if X.close(moved, 0) else ""
            wrong.append(f"'{label}' changed by {_amt(moved, sign=True)}, expected {_amt(exp, sign=True)}{same}")
    require(not wrong, "in a copy of the file where every existing expense got a different amount, category and "
                       f"month, {len(wrong)} of {len(refs)} Summary totals did not follow: " + "; ".join(wrong[:3])
                       + (f" (and {len(wrong) - 3} more)" if len(wrong) > 3 else ""))


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
