"""Alternative valid solution: SUMIF over the rows that have a rep name (subtotal rows have none)."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "regional.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Sales"]
    last = ws.max_row
    total = last + 1
    ws.cell(total, 1, "Grand Total")
    for col in "CDEFG":
        ws[f"{col}{total}"] = f'=SUMIF($B$2:$B${last},"<>",{col}2:{col}{last})'
    wb.save(path)
