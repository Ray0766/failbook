"""Reproduces the failure: compute the margins in Python and write them as numbers."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "pricing.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Products"]
    ws.cell(1, 5, "Margin %")
    ws.cell(1, 5)._style = ws.cell(1, 4)._style
    for r in range(2, ws.max_row + 1):
        cost, price = ws.cell(r, 3).value, ws.cell(r, 4).value
        if price:
            ws.cell(r, 5, (price - cost) / price).number_format = "0.0%"
    wb.save(path)
