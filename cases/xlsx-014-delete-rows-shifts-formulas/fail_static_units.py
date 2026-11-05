"""Wrong: rows are deleted and the formulas rewritten, but Summary 'Total units' is typed in as a number."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "inventory.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Stock"]
    for r in range(ws.max_row, 1, -1):
        if ws.cell(r, 3).value == "Discontinued":
            ws.delete_rows(r)
    total = next(r for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value == "Total")
    last = total - 1
    for r in range(2, total):
        ws.cell(r, 6).value = f"=D{r}*E{r}"
    ws.cell(total, 4).value = f"=SUM(D2:D{last})"
    ws.cell(total, 6).value = f"=SUM(F2:F{last})"
    summary = wb["Summary"]
    summary["B1"] = f'=COUNTIF(Stock!C2:C{last},"Active")'
    summary["B2"] = sum(it["qty"] for it in meta["items"] if it["status"] != "Discontinued")
    summary["B3"] = f"=Stock!F{total}"
    wb.save(path)
