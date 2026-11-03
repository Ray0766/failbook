"""Wrong: the typo is fixed, but an autofilter is added to the header row."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "customers.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Customers"]
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 3).value == "Bejing":
            ws.cell(r, 3).value = "Beijing"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)
