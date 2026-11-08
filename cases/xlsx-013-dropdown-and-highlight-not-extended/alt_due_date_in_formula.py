"""Alternative valid solution: the highlight formula also requires a due date (AND($D2="Blocked",$E2>0)); every task
has one, so exactly the Blocked tasks are red. Dates compare as Excel serial numbers."""

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.cell_range import MultiCellRange

from failbook import kit


def solve(workdir, meta):
    path = workdir / "tracker.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Tracker"]
    for t in meta["new"]:
        ws.append([t["id"], t["task"], t["owner"], t["status"], kit.from_iso(t["due"])])
        ws.cell(ws.max_row, 5).number_format = kit.DATE_FMT
    last = ws.max_row
    for dv in ws.data_validations.dataValidation:
        dv.sqref = MultiCellRange(f"D2:D{last}")
    ws.conditional_formatting = ConditionalFormattingList()
    ws.conditional_formatting.add(
        f"D2:D{last}",
        FormulaRule(formula=['AND($D2="Blocked",$E2>0)'], fill=PatternFill(bgColor="FFFFC7CE"),
                    font=Font(color="FF9C0006")),
    )
    wb.save(path)
