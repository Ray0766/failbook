"""Wrong: the Status validation is replaced by a new one with openpyxl's defaults, which turn the error alert off,
so every Status cell accepts any text."""

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.worksheet.datavalidation import DataValidation, DataValidationList

from failbook import kit


def solve(workdir, meta):
    path = workdir / "tracker.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Tracker"]
    for t in meta["new"]:
        ws.append([t["id"], t["task"], t["owner"], t["status"], kit.from_iso(t["due"])])
        ws.cell(ws.max_row, 5).number_format = kit.DATE_FMT
    last = ws.max_row
    ws.data_validations = DataValidationList()
    dv = DataValidation(type="list", formula1='"Not started,In progress,Blocked,Done"', allow_blank=True)
    dv.add(f"D2:D{last}")
    ws.add_data_validation(dv)
    old = ws.conditional_formatting
    ws.conditional_formatting = ConditionalFormattingList()
    for cf in old:
        for cf_rule in cf.rules:
            ws.conditional_formatting.add(f"D2:D{last}", cf_rule)
    wb.save(path)
