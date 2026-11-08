"""Reproduces the failure: sort the rows in Python and write each cell back as-is (formulas as text)."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "staff.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Staff"]
    rows = [[c.value for c in row] for row in ws.iter_rows(min_row=2, max_row=ws.max_row)]
    rows.sort(key=lambda row: (row[1], row[0]))
    for r, row in enumerate(rows, start=2):
        for c, value in enumerate(row, start=1):
            ws.cell(r, c).value = value
    wb.save(path)
