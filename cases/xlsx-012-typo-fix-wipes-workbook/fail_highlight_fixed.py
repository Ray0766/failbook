"""Wrong: the typo is fixed, but each corrected cell is highlighted (yellow fill, bold red font)."""

import openpyxl
from openpyxl.styles import Font, PatternFill


def solve(workdir, meta):
    path = workdir / "customers.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Customers"]
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 3).value == "Bejing":
            ws.cell(r, 3).value = "Beijing"
            ws.cell(r, 3).fill = PatternFill("solid", fgColor="FFFF00")
            ws.cell(r, 3).font = Font(bold=True, color="C00000")
    wb.save(path)
