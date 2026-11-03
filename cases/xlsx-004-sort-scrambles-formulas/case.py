from collections import Counter

import openpyxl

from failbook import kit
from failbook import xlsx as X
from failbook.checks import require, rule

SHEET = "Staff"
HEADERS = ["Name", "Department", "Base Salary", "Bonus Rate", "Bonus"]


def params(seed):
    r = kit.rng(seed, "xlsx-004")
    n = r.randint(14, 18)
    names = kit.people(r, n)
    staff = [
        {
            "name": name,
            "dept": r.choice(kit.DEPARTMENTS[:4]),
            "salary": r.randrange(42000, 120000, 500),
            "rate": r.choice([0.03, 0.05, 0.08, 0.1, 0.12, 0.15]),
        }
        for name in names
    ]
    # Make sure sorting actually moves the first row.
    ordered = sorted(staff, key=lambda s: (s["dept"].casefold(), s["name"].casefold()))
    if ordered[0] is staff[0]:
        staff.append(staff.pop(0))
    return {"staff": staff}


def build(meta, outdir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.append(HEADERS)
    kit.style_header(ws, 1, len(HEADERS))
    for r, s in enumerate(meta["staff"], start=2):
        ws.append([s["name"], s["dept"], s["salary"], s["rate"], f"=C{r}*D{r}"])
        ws.cell(r, 3).number_format = kit.INT_FMT
        ws.cell(r, 4).number_format = "0%"
        ws.cell(r, 5).number_format = kit.INT_FMT
    kit.set_widths(ws, {"A": 18, "B": 14, "C": 12, "D": 11, "E": 10})
    ws.freeze_panes = "A2"
    wb.save(outdir / "staff.xlsx")


def _rows(ctx):
    return range(2, 2 + len(ctx.meta["staff"]))


def _records(ctx, values=None):
    vals = (values or ctx.values()).ws(SHEET)
    out = []
    for r in _rows(ctx):
        out.append((r, [vals.cell(r, c).value for c in range(1, 5)]))
    return out


def _by_name(ctx):
    return {s["name"]: s for s in ctx.meta["staff"]}


def scenarios(ctx):
    ws = ctx.ws(SHEET)
    by_name = _by_name(ctx)
    overrides = {}
    for r in _rows(ctx):
        person = by_name.get(ws.cell(r, 1).value)
        if person:
            overrides[X.make_ref(SHEET, (r, 3))] = person["salary"] + 1000 * (r - 1)
    return {"raise": overrides}


@rule(
    "header-kept",
    "The header row is still at the top",
    kind="preserve",
    title_zh="表头仍在第一行",
    requirement="Keep the header row at the top.",
)
def header_kept(ctx):
    ws = ctx.ws(SHEET)
    got = [ws.cell(1, c).value for c in range(1, 6)]
    require([X.norm_text(v) for v in got] == [X.norm_text(h) for h in HEADERS], f"row 1 is {got}")


@rule(
    "same-people",
    "Every person's row is intact: nobody lost, duplicated or mixed up",
    kind="preserve",
    title_zh="每个人的行完整：无丢失、无重复、无串行",
    requirement="Keep every person's data together on one row; nobody may be lost or duplicated.",
)
def same_people(ctx):
    expected = Counter((s["name"], s["dept"], s["salary"], s["rate"]) for s in ctx.meta["staff"])
    got = Counter()
    for _, rec in _records(ctx):
        name, dept, salary, rate = rec
        got[(name, dept, salary, round(rate, 6) if X.is_number(rate) else rate)] += 1
    expected = Counter({(n, d, s, round(r, 6)): k for (n, d, s, r), k in expected.items()})
    missing = expected - got
    extra = got - expected
    require(not missing and not extra,
            f"rows differ from the original: missing {sorted(missing)[:3]}, unexpected {sorted(extra)[:3]}")
    vals = ctx.values().ws(SHEET)
    stray = vals.cell(2 + len(ctx.meta["staff"]), 1).value
    require(stray in (None, ""), f"there is an extra row below the table: {X.show(stray)}")


@rule(
    "sorted",
    "Rows are sorted by Department, then by Name",
    kind="content",
    title_zh="已按部门、再按姓名排序",
    requirement="Sort the table by Department (A-Z), and within each department by Name (A-Z).",
)
def sorted_rows(ctx):
    keys = [(X.norm_text(rec[1]), X.norm_text(rec[0])) for _, rec in _records(ctx)]
    for i in range(1, len(keys)):
        require(keys[i - 1] <= keys[i], f"rows {i + 1} and {i + 2} are out of order: {keys[i - 1]} then {keys[i]}")


@rule(
    "bonus-matches-row",
    "Each Bonus equals Base Salary x Bonus Rate of the same row",
    kind="value",
    title_zh="每行的 Bonus 等于同一行的 Base Salary × Bonus Rate",
    requirement="The Bonus column must remain a formula: Base Salary x Bonus Rate of the same row.",
)
def bonus_matches_row(ctx):
    values = ctx.values()
    by_name = _by_name(ctx)
    wrong = []
    for r, rec in _records(ctx, values):
        person = by_name.get(rec[0])
        if not person:
            continue
        got = ctx.num(values, X.make_ref(SHEET, (r, 5)), f"Bonus of {rec[0]}")
        exp = person["salary"] * person["rate"]
        if not X.close(got, exp):
            wrong.append(f"{rec[0]} (row {r}): {got:,.0f} instead of {exp:,.0f}")
    require(not wrong, f"{len(wrong)} bonus(es) belong to another row, e.g. " + "; ".join(wrong[:3]))


@rule(
    "bonus-is-live",
    "Changing a salary changes the Bonus on that same row",
    kind="live",
    title_zh="修改某人的薪资后，同一行的 Bonus 随之变化",
    requirement="The Bonus column must remain a formula: Base Salary x Bonus Rate of the same row.",
)
def bonus_is_live(ctx):
    values = ctx.values(scenario="raise")
    by_name = _by_name(ctx)
    ws = ctx.ws(SHEET)
    for r in _rows(ctx):
        person = by_name.get(ws.cell(r, 1).value)
        if not person:
            continue
        salary = person["salary"] + 1000 * (r - 1)
        got = ctx.num(values, X.make_ref(SHEET, (r, 5)), f"Bonus of {person['name']}")
        require(X.close(got, salary * person["rate"]),
                f"after changing {person['name']}'s salary, their Bonus is {got:,.0f}, expected {salary * person['rate']:,.0f}")
