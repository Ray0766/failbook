import datetime as dt

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from failbook import kit
from failbook import xlsx as X
from failbook.checks import Inconclusive, require, rule

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
    """The options a list dropdown offers: (options, None), or (None, why) if its source is not resolved to cells."""
    f = (dv.formula1 or "").strip().lstrip("=")
    if f.startswith('"') and f.endswith('"'):
        return [o.strip() for o in f[1:-1].split(",")], None
    rng = X.resolve_range(f, ctx.wb(), SHEET)
    if rng is None:
        return None, "which the checker cannot resolve to fixed cells"
    sheet, r1, c1, r2, c2 = rng
    calc = ctx.values().wb
    real = X.sheet_named(calc, sheet)  # sheet names match in any letter case, as in Excel and LibreOffice
    if real is None:
        return None, f"a range on a sheet {sheet!r} that the workbook does not have"
    ws = calc[real]
    cells = (ws.cell(r, c).value for r in range(r1, r2 + 1) for c in range(c1, c2 + 1))
    return [v for v in cells if v not in (None, "")], None


def _listing(cells, limit=6):
    cells = list(cells)
    more = f" and {len(cells) - limit} more" if len(cells) > limit else ""
    return ", ".join(cells[:limit]) + more


def _rejects_other_entries(dv):
    """Whether a validation refuses an entry that is not in its list: an error alert of style "stop" (the default)."""
    return bool(dv.showErrorMessage) and (dv.errorStyle or "stop") == "stop"


def _alert_problem(dv):
    if not dv.showErrorMessage:
        return f"the validation on {dv.sqref} has its error alert off"
    style = dv.errorStyle or "stop"
    article = "an" if style[:1].lower() in "aeiou" else "a"
    return f"the validation on {dv.sqref} has {article} {style!r} error alert, which lets the entry stand"


# -- what the conditional formats highlight -------------------------------------------
def _cf_rules(ctx, ws):
    """Every conditional-format rule on the sheet as (origin, rule, sqref, colours), highest priority first.

    ``origin`` is the top-left cell of the rule's range: relative references in its
    formula are written for that cell. ``colours`` is what its style sets (``X.dxf_colours``).
    """
    patterns = X.dxf_fill_patterns(ctx.path())
    wb = ctx.wb()
    out = []
    for cf in ws.conditional_formatting:
        ranges = list(cf.sqref.ranges)
        if not ranges:
            continue
        origin = (min(r.min_row for r in ranges), min(r.min_col for r in ranges))
        for cf_rule in cf.rules:
            i = cf_rule.dxfId
            raw = patterns[i] if isinstance(i, int) and 0 <= i < len(patterns) else None
            colours = X.dxf_colours(cf_rule.dxf, wb, raw_pattern=raw)
            out.append((cf_rule.priority if cf_rule.priority is not None else 10**9, origin, cf_rule, cf.sqref,
                        colours))
    out.sort(key=lambda t: t[0])
    return [t[1:] for t in out]


def _value_at(calc_wb, override=None):
    """``value_at(sheet, row, col)`` over the recalculated workbook for the evaluator.

    Sheet names match case-insensitively, as in Excel. ``override`` is ``((row, col), value)``:
    that Tracker cell reads as ``value`` (to try each status).
    """
    def get(sheet, row, col):
        real = X.sheet_named(calc_wb, sheet)
        if real is None:
            raise X.Unknown(f"the formula refers to a sheet {sheet!r} that the workbook does not have")
        if override is not None and real == SHEET and (row, col) == override[0]:
            return override[1]
        return calc_wb[real].cell(row, col).value
    return get


def _describe_rule(cf_rule, sqref):
    if cf_rule.type == "cellIs":
        what = f"cell value {cf_rule.operator} {', '.join(map(str, cf_rule.formula or []))}"
    elif cf_rule.formula:
        what = f"{cf_rule.type} {cf_rule.formula[0]}"
    else:
        what = f"{cf_rule.type} {cf_rule.text or ''}".strip()
    return f"the rule '{what}' on {sqref}"


def _red_at(ctx, rules, cell, value_at):
    """Whether conditional formatting shows ``cell`` in red: (True | False | None, reason)."""
    row, col = X.coord_to_rc(cell)
    wb = ctx.wb()
    decided = {}  # "fill"/"font" -> colour of the highest-priority rule that applies and sets it
    not_applied = []  # why each rule with a colour that covers the cell does not apply there
    stopped = None  # the rule with "stop if true" that applies there and ends the evaluation
    for origin, cf_rule, sqref, sets in rules:
        if cell not in sqref:
            continue
        if not sets and not cf_rule.stopIfTrue:
            continue
        try:
            applies, note = X.cf_rule_outcome(cf_rule, SHEET, origin, (row, col), value_at, wb)
        except X.Unknown as exc:
            if cf_rule.stopIfTrue or any(k not in decided for k in sets):
                return None, f"the checker cannot evaluate {_describe_rule(cf_rule, sqref)}: {exc}"
            continue
        if not applies:
            not_applied.append(f"{_describe_rule(cf_rule, sqref)} {note or 'is false there'}")
            continue
        for key, colour in sets.items():
            decided.setdefault(key, colour)
        if cf_rule.stopIfTrue:
            stopped = f"{_describe_rule(cf_rule, sqref)} (priority {cf_rule.priority}, stop if true) applies there"
            break
    if not decided:
        if stopped:
            return False, f"{stopped} and sets no colour, so the rules after it are not applied"
        if not not_applied:
            return False, "no conditional format that sets a colour applies to it"
        return False, " and ".join(not_applied[:2])
    reds = {key: None if colour == "?" else X.is_red(colour) for key, colour in decided.items()}
    shown = ", ".join(f"{key} {colour}" for key, colour in decided.items())
    if any(reds.values()):
        return True, f"it is shown with {shown}"
    if all(r is False for r in reds.values()):
        after = f"; {stopped}, so the rules after it are not applied" if stopped else ""
        return False, f"it is highlighted with {shown}, which is not red{after}"
    unsure = []
    for key, colour in decided.items():
        if reds[key] is not None:
            continue
        if colour == "?":
            unsure.append(getattr(colour, "why", None) or f"its {key} colour is not resolved by the checker")
        else:
            unsure.append(f"its {key} {colour} is a pale red (a red hue with saturation from {X.PALE_SATURATION:.0%} "
                          f"to below {X.RED_SATURATION:.0%} and brightness at least {X.PALE_BRIGHTNESS:.0%}), which "
                          "the checker does not judge as red or not red")
    return None, "; ".join(unsure)


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
    requirement="The new rows must work like the existing ones: the Status dropdown ... must apply to them too.",
)
def dropdown_covers_new_rows(ctx):
    ws = ctx.ws(SHEET)
    wanted = {X.norm_text(s) for s in STATUSES}
    # cell -> [(offers the four statuses: True / False / None if unknown, validation, its options)]
    found = {c: [] for c in _status_cells(ctx)}
    unresolved = {}  # id(validation) -> why its options are not known
    for dv in ws.data_validations.dataValidation:
        if dv.type != "list":
            continue
        options, why = _list_options(ctx, dv)
        ok = None if options is None else {X.norm_text(o) for o in options} == wanted
        if why:
            unresolved[id(dv)] = why
        for cell, got in found.items():
            if cell in dv.sqref:
                got.append((ok, dv, options))
    uncovered = [c for c, got in found.items() if not got]
    require(not uncovered, f"the Status dropdown does not cover {_listing(uncovered)}")
    # showDropDown="1" (openpyxl: showDropDown=True, alias hide_drop_down) hides the in-cell list.
    hidden = [c for c, got in found.items() if all(dv.showDropDown for _, dv, _ in got)]
    if hidden:
        dv = found[hidden[0]][0][1]
        require(False, f"{_listing(hidden)} {'has' if len(hidden) == 1 else 'have'} no dropdown: the list "
                f"validation on {dv.sqref} hides the in-cell list (showDropDown=\"1\")")
    found = {c: [g for g in got if not g[1].showDropDown] for c, got in found.items()}
    wrong = [(c, got) for c, got in found.items() if all(ok is False for ok, _, _ in got)]
    if wrong:
        options = wrong[0][1][0][2]
        require(False, f"the dropdown on {_listing(c for c, _ in wrong)} offers {options}, not the four statuses")
    original = [dv for dv in ctx.input_wb()[SHEET].data_validations.dataValidation
                if dv.type == "list" and "D2" in dv.sqref]
    if any(_rejects_other_entries(dv) for dv in original):
        lax = [c for c, got in found.items() if not any(_rejects_other_entries(dv) for _, dv, _ in got)]
        if lax:
            require(False, f"{_listing(lax)} {'accepts' if len(lax) == 1 else 'accept'} any text: "
                    f"{_alert_problem(found[lax[0]][0][1])}, unlike the original")
    unknown = [(c, got) for c, got in found.items() if not any(ok for ok, _, _ in got)]
    if unknown:
        dv = next(dv for ok, dv, _ in unknown[0][1] if ok is None)
        raise Inconclusive(
            f"the dropdown on {_listing(c for c, _ in unknown)} takes its options from {dv.formula1}, "
            f"{unresolved[id(dv)]}; whether it offers the four statuses is not checked")


@rule(
    "highlight-covers-new-rows",
    "Every task row, including the new ones, is highlighted in red exactly when its Status is 'Blocked'",
    kind="structure",
    title_zh="所有任务行（含新行）在且仅在 Status 为 'Blocked' 时显示红色高亮",
    requirement="A Status of \"Blocked\" is highlighted in red ... the red highlighting for \"Blocked\" must apply to "
                "them too.",
)
def highlight_covers_new_rows(ctx):
    ws = ctx.ws(SHEET)
    calc = ctx.values()
    rules = _cf_rules(ctx, ws)

    now_missing, now_wrong, if_missing, if_wrong, unknown = [], [], [], [], []
    for cell in _status_cells(ctx):
        rc = X.coord_to_rc(cell)
        current = calc.wb[SHEET].cell(*rc).value
        for status in [current] + [s for s in STATUSES if s != current]:
            red, why = _red_at(ctx, rules, cell, _value_at(calc.wb, (rc, status)))
            blocked = X.norm_text(status) == "blocked"
            if red is None:
                unknown.append((cell, status, why))
            elif blocked and not red:
                (now_missing if status == current else if_missing).append((cell, current, status, why))
            elif red and not blocked:
                (now_wrong if status == current else if_wrong).append((cell, current, status, why))
    if now_missing:
        cell, _, _, why = now_missing[0]
        require(False, f"{_listing(m[0] for m in now_missing)} (Status 'Blocked') "
                f"{'is' if len(now_missing) == 1 else 'are'} not highlighted in red; for {cell}: {why}")
    if now_wrong:
        cell, current, _, why = now_wrong[0]
        require(False, f"{_listing(m[0] for m in now_wrong)} {'is' if len(now_wrong) == 1 else 'are'} highlighted "
                f"in red although the Status is not 'Blocked' ({cell} is {current!r}; {why})")
    if if_missing:
        cell, current, _, why = if_missing[0]
        require(False, f"{_listing(m[0] for m in if_missing)} would not turn red if the Status were changed to "
                f"'Blocked' ({cell} is now {current!r}; {why})")
    if if_wrong:
        cell, current, status, why = if_wrong[0]
        require(False, f"{_listing(m[0] for m in if_wrong)} would be highlighted in red with a Status other than "
                f"'Blocked' ({cell} set to {status!r}; {why})")
    if unknown:
        cell, status, why = unknown[0]
        raise Inconclusive(f"cannot tell whether {cell} is highlighted in red when its Status is {status!r}: {why}")


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
