"""Wrong: the new rows get their own list validation with showDropDown=True, which hides the in-cell list."""

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.worksheet.datavalidation import DataValidation

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
    # In SpreadsheetML showDropDown="1" means "hide the dropdown" (openpyxl's alias is hide_drop_down).
    dv = DataValidation(type="list", formula1='"Not started,In progress,Blocked,Done"', allow_blank=True,
                        showDropDown=True, showErrorMessage=True, error="Pick a status from the list",
                        errorTitle="Invalid status")
    dv.add(f"D{first}:D{last}")
    ws.add_data_validation(dv)
    old = ws.conditional_formatting
    ws.conditional_formatting = ConditionalFormattingList()
    for cf in old:
        for cf_rule in cf.rules:
            ws.conditional_formatting.add(f"D2:D{last}", cf_rule)
    wb.save(path)
