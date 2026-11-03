"""Reproduces the failure: append the rows and assume the chart follows the data."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "monthly.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Monthly"]
    for m in meta["new"]:
        ws.append([m["month"], m["revenue"], m["costs"]])
        ws.cell(ws.max_row, 2).number_format = ws.cell(ws.max_row, 3).number_format = "#,##0"
    wb.save(path)
