"""Reproduces the failure: openpyxl delete_rows(), bottom-up, without rewriting formulas."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "inventory.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Stock"]
    for r in range(ws.max_row, 1, -1):
        if ws.cell(r, 3).value == "Discontinued":
            ws.delete_rows(r)
    wb.save(path)
