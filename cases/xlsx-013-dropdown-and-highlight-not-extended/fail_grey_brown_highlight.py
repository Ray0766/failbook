"""Wrong: the "Blocked" highlight covers every row, but its fill is a dark greyish brown (5A5050), not red."""

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import CellIsRule
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
    last = ws.max_row
    for dv in ws.data_validations.dataValidation:
        dv.sqref = MultiCellRange(f"D2:D{last}")
    ws.conditional_formatting = ConditionalFormattingList()
    ws.conditional_formatting.add(
        f"D2:D{last}",
        CellIsRule(operator="equal", formula=['"Blocked"'], fill=PatternFill(bgColor="FF5A5050")),
    )
    wb.save(path)
