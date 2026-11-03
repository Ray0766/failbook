"""Reproduces the failure: SUM the whole column, subtotal rows included."""

import openpyxl
from openpyxl.styles import Font


def solve(workdir, meta):
    path = workdir / "regional.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Sales"]
    last = ws.max_row
    total = last + 1
    ws.cell(total, 1, "Grand Total").font = Font(bold=True)
    for col in "CDEFG":
        cell = ws[f"{col}{total}"]
        cell.value = f"=SUM({col}2:{col}{last})"
        cell.font = Font(bold=True)
        cell.number_format = "#,##0"
    wb.save(path)
