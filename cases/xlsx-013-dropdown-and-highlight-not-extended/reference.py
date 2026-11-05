"""A correct solution: append the rows and extend the dropdown and the highlight to them."""

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList

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
        if "D2" in dv.sqref:
            dv.sqref = openpyxl.worksheet.cell_range.MultiCellRange(f"D2:D{last}")
    old = ws.conditional_formatting
    ws.conditional_formatting = ConditionalFormattingList()
    for cf in old:
        target = f"D2:D{last}" if "D2" in cf.sqref else str(cf.sqref)
        for cf_rule in cf.rules:
            ws.conditional_formatting.add(target, cf_rule)
    wb.save(path)
