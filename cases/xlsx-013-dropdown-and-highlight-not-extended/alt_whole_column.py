"""Alternative valid solution: apply the dropdown to the whole column and highlight with a formula rule."""

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.worksheet.cell_range import MultiCellRange

from failbook import kit


def solve(workdir, meta):
    path = workdir / "tracker.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Tracker"]
    for t in meta["new"]:
        ws.append([t["id"], t["task"], t["owner"], t["status"], kit.from_iso(t["due"])])
        ws.cell(ws.max_row, 5).number_format = kit.DATE_FMT
    for dv in ws.data_validations.dataValidation:
        dv.sqref = MultiCellRange("D2:D1000")
    ws.conditional_formatting = ConditionalFormattingList()
    ws.conditional_formatting.add(
        "D2:D1000",
        FormulaRule(formula=['$D2="Blocked"'], fill=PatternFill("solid", start_color="FFFF0000", end_color="FFFF0000")),
    )
    wb.save(path)
