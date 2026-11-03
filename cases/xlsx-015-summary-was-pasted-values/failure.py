"""Reproduces the failure: correct the data cells and assume the summary recalculates."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "budget.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Q3 Data"]
    fixes = {(c["dept"], c["month"]): c["new"] for c in meta["corrections"]}
    for r in range(2, ws.max_row + 1):
        key = (ws.cell(r, 1).value, ws.cell(r, 2).value)
        if key in fixes:
            ws.cell(r, 3).value = fixes[key]
    wb.save(path)
