"""Wrong: the old rule is kept, and the new rows get a copy of its formula written for D2, so they follow rows 2.. ."""

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.worksheet.cell_range import MultiCellRange

from failbook import kit


def solve(workdir, meta):
    path = workdir / "tracker.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Tracker"]
    first = ws.max_row + 1
    for t in meta["new"]:
        ws.append([t["id"], t["task"], t["owner"], t["status"], kit.from_iso(t["due"])])
        ws.cell(ws.max_row, 5).number_format = kit.DATE_FMT
    last = ws.max_row
    for dv in ws.data_validations.dataValidation:
        dv.sqref = MultiCellRange(f"D2:D{last}")
    ws.conditional_formatting.add(
        f"D{first}:D{last}",
        FormulaRule(formula=['D2="Blocked"'], fill=PatternFill("solid", start_color="FFFFC7CE", end_color="FFFFC7CE")),
    )
    wb.save(path)
