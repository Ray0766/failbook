"""Alternative valid solution: the dropdown takes its options from a defined name on a Lists sheet."""

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.cell_range import MultiCellRange

from failbook import kit

STATUSES = ["Not started", "In progress", "Blocked", "Done"]


def solve(workdir, meta):
    path = workdir / "tracker.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Tracker"]
    for t in meta["new"]:
        ws.append([t["id"], t["task"], t["owner"], t["status"], kit.from_iso(t["due"])])
        ws.cell(ws.max_row, 5).number_format = kit.DATE_FMT
    last = ws.max_row
    lists = wb.create_sheet("Lists")
    for i, status in enumerate(STATUSES, start=1):
        lists.cell(i, 1, status)
    wb.defined_names["StatusList"] = DefinedName("StatusList", attr_text=f"Lists!$A$1:$A${len(STATUSES)}")
    for dv in ws.data_validations.dataValidation:
        dv.sqref = MultiCellRange(f"D2:D{last}")
        dv.formula1 = "StatusList"
    old = ws.conditional_formatting
    ws.conditional_formatting = ConditionalFormattingList()
    for cf in old:
        for cf_rule in cf.rules:
            ws.conditional_formatting.add(f"D2:D{last}", cf_rule)
    wb.save(path)
