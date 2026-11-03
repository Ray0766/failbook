import datetime as dt

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from failbook import kit
from failbook import xlsx as X
from failbook.checks import require, rule

SHEET = "Tracker"
HEADERS = ["ID", "Task", "Owner", "Status", "Due"]
STATUSES = ["Not started", "In progress", "Blocked", "Done"]
TASKS = [
    "Update price list", "Review supplier contract", "Prepare Q3 report", "Fix invoice template",
    "Plan team offsite", "Migrate CRM data", "Audit expense claims", "Draft hiring plan",
    "Renew software licences", "Clean up shared drive", "Write onboarding guide", "Order new laptops",
    "Check warehouse stock", "Book trade fair stand", "Translate product sheet", "Archive old projects",
    "Set up backup job", "Collect customer feedback", "Update org chart", "Test new POS terminal",
    "Prepare budget draft", "Survey office chairs", "Send tax documents", "Organise training day",
]


def params(seed):
    r = kit.rng(seed, "xlsx-013")
    n = r.randint(12, 18)
    owners = kit.people(r, 5)
    names = r.sample(TASKS, n + 5)
    start = dt.date(2024, 6, 3)
    tasks = []
    for i, name in enumerate(names):
        tasks.append({
            "id": f"T-{101 + i}",
            "task": name,
            "owner": r.choice(owners),
            "status": r.choice(STATUSES),
            "due": (start + dt.timedelta(days=r.randint(0, 60))).isoformat(),
        })
    new = tasks[n:]
    new[1]["status"] = "Blocked"
    table = kit.md_table(HEADERS, [(t["id"], t["task"], t["owner"], t["status"], t["due"]) for t in new])
    return {"existing": tasks[:n], "new": new, "new_md": table}


def build(meta, outdir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.append(HEADERS)
    kit.style_header(ws, 1, len(HEADERS))
    for t in meta["existing"]:
        ws.append([t["id"], t["task"], t["owner"], t["status"], kit.from_iso(t["due"])])
        ws.cell(ws.max_row, 5).number_format = kit.DATE_FMT
    last = ws.max_row
    dv = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=True)
    dv.error, dv.errorTitle, dv.showErrorMessage = "Pick a status from the list", "Invalid status", True
    dv.add(f"D2:D{last}")
    ws.add_data_validation(dv)
    red = CellIsRule(operator="equal", formula=['"Blocked"'],
                     fill=PatternFill("solid", start_color="FFFFC7CE", end_color="FFFFC7CE"),
                     font=Font(color="FF9C0006", bold=True))
    ws.conditional_formatting.add(f"D2:D{last}", red)
    kit.set_widths(ws, {"A": 8, "B": 28, "C": 16, "D": 13, "E": 12})
    ws.freeze_panes = "A2"
    wb.save(outdir / "tracker.xlsx")


def _status_cells(ctx):
    n = len(ctx.meta["existing"]) + len(ctx.meta["new"])
    return [f"D{r}" for r in range(2, 2 + n)]


def _new_cells(ctx):
    first = 2 + len(ctx.meta["existing"])
    return [f"D{r}" for r in range(first, first + len(ctx.meta["new"]))]


def _list_options(ctx, dv):
    f = (dv.formula1 or "").strip()
    if f.startswith('"'):
        return [o.strip() for o in f.strip('"').split(",")]
    rng = X.parse_range(f.lstrip("="))
    if rng is None:
        return []
    sheet, r1, c1, r2, c2 = rng
    vals = ctx.values().ws(sheet)
    return [vals.cell(r, c).value for r in range(r1, r2 + 1) for c in range(c1, c2 + 1) if vals.cell(r, c).value]


def _is_blocked_rule(cf_rule):
    texts = [str(f) for f in (cf_rule.formula or [])] + [str(getattr(cf_rule, "text", "") or "")]
    mentions = any("blocked" in t.casefold() for t in texts)
    styled = cf_rule.dxf is not None and (cf_rule.dxf.fill is not None or cf_rule.dxf.font is not None)
    kind_ok = cf_rule.type in ("cellIs", "expression", "containsText")
    if cf_rule.type == "cellIs":
        kind_ok = cf_rule.operator in ("equal", None)
    return mentions and styled and kind_ok


@rule(
    "tasks-added",
    "The new tasks are appended after the existing ones with the given values",
    kind="content",
    title_zh="新任务按给定内容追加在现有任务之后",
    requirement="Add these tasks at the end of the tracker.",
)
def tasks_added(ctx):
    vals = ctx.values().ws(SHEET)
    first = 2 + len(ctx.meta["existing"])
    for i, t in enumerate(ctx.meta["new"]):
        r = first + i
        got = [vals.cell(r, c).value for c in range(1, 6)]
        require(got[:4] == [t["id"], t["task"], t["owner"], t["status"]], f"row {r} is {got[:4]}, expected {t['id']} ...")
        require(X.as_date(got[4]) == kit.from_iso(t["due"]), f"Due of {t['id']} is {X.show(got[4])}, expected {t['due']}")


@rule(
    "dropdown-covers-new-rows",
    "The Status dropdown applies to every task row, including the new ones",
    kind="structure",
    title_zh="Status 下拉列表覆盖所有任务行（含新行）",
    requirement="The Status dropdown ... must apply to them too.",
)
def dropdown_covers_new_rows(ctx):
    ws = ctx.ws(SHEET)
    wanted = {X.norm_text(s) for s in STATUSES}
    covered = set()
    for dv in ws.data_validations.dataValidation:
        if dv.type == "list" and {X.norm_text(o) for o in _list_options(ctx, dv)} == wanted:
            covered |= {c for c in _status_cells(ctx) if c in dv.sqref}
    missing = [c for c in _status_cells(ctx) if c not in covered]
    require(not missing, f"the Status dropdown does not cover {', '.join(missing)}")


@rule(
    "highlight-covers-new-rows",
    "The red 'Blocked' highlight applies to every task row, including the new ones",
    kind="structure",
    title_zh="'Blocked' 红色高亮覆盖所有任务行（含新行）",
    requirement="The red highlighting for 'Blocked' must apply to them too.",
)
def highlight_covers_new_rows(ctx):
    ws = ctx.ws(SHEET)
    covered = set()
    for cf in ws.conditional_formatting:
        if any(_is_blocked_rule(r) for r in cf.rules):
            covered |= {c for c in _status_cells(ctx) if c in cf.sqref}
    missing = [c for c in _status_cells(ctx) if c not in covered]
    require(not missing, f"the 'Blocked' highlight does not cover {', '.join(missing)}")


@rule(
    "existing-unchanged",
    "The existing tasks are unchanged",
    kind="preserve",
    title_zh="原有任务保持不变",
    requirement="Do not change anything the task does not ask you to change.",
)
def existing_unchanged(ctx):
    vals = ctx.values().ws(SHEET)
    for r, t in enumerate(ctx.meta["existing"], start=2):
        got = [vals.cell(r, c).value for c in range(1, 5)]
        require(got == [t["id"], t["task"], t["owner"], t["status"]], f"row {r} changed: {got}")
