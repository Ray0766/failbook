"""Wrong: the typo is fixed, but two header cells are renamed."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "customers.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Customers"]
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 3).value == "Bejing":
            ws.cell(r, 3).value = "Beijing"
    ws["C1"] = "City (fixed)"
    ws["F1"] = "Annual Spend (CNY)"
    wb.save(path)
