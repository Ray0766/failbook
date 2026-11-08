"""Alternative valid solution: extend the dropdown and highlight with a red "text begins with 'Blocked'" rule."""

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import Rule
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.differential import DifferentialStyle
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
    style = DifferentialStyle(fill=PatternFill(bgColor="FFFFC7CE"), font=Font(color="FF9C0006", bold=True))
    ws.conditional_formatting.add(
        f"D2:D{last}",
        Rule(type="beginsWith", operator="beginsWith", text="Blocked", dxf=style,
             formula=['LEFT(D2,LEN("Blocked"))="Blocked"']),
    )
    wb.save(path)
