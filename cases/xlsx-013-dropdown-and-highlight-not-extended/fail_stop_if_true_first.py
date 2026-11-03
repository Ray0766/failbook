"""Wrong: a bold-only "Blocked" rule with Stop If True comes before the red one, so the red rule never applies."""

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import CellIsRule
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
        f"D2:D{last}", CellIsRule(operator="equal", formula=['"Blocked"'], font=Font(bold=True), stopIfTrue=True))
    ws.conditional_formatting.add(
        f"D2:D{last}", CellIsRule(operator="equal", formula=['"Blocked"'], fill=PatternFill(bgColor="FFFFC7CE")))
    wb.save(path)
