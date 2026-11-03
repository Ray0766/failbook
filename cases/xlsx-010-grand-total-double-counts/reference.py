"""A correct solution: add up the subtotal rows only (each one already covers its region)."""

import openpyxl
from openpyxl.styles import Font


def solve(workdir, meta):
    path = workdir / "regional.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Sales"]
    subtotal_rows = [r for r in range(2, ws.max_row + 1)
                     if str(ws.cell(r, 1).value or "").endswith("subtotal")]
    total = ws.max_row + 1
    ws.cell(total, 1, "Grand Total").font = Font(bold=True)
    for col in "CDEFG":
        cell = ws[f"{col}{total}"]
        cell.value = "=" + "+".join(f"{col}{r}" for r in subtotal_rows)
        cell.font = Font(bold=True)
        cell.number_format = "#,##0"
    wb.save(path)
