"""A wrong solution: the Total formulas are extended correctly, but the Amount of the first two new
orders is typed in as a number (Qty x Unit Price computed outside the sheet) instead of a formula.
Today's values are right; the task asks for Amount "with a formula", so a later change of Qty or
Unit Price would leave these two Amounts and the Total Amount stale.
"""

import openpyxl

from failbook import kit


def solve(workdir, meta):
    path = workdir / "sales.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Orders"]
    total_row = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "Total")
    n = len(meta["new"])
    ws.insert_rows(total_row, amount=n)
    for i, o in enumerate(meta["new"]):
        r = total_row + i
        ws.cell(r, 1, o["id"])
        ws.cell(r, 2, kit.from_iso(o["date"])).number_format = kit.DATE_FMT
        ws.cell(r, 3, o["customer"])
        ws.cell(r, 4, o["qty"])
        ws.cell(r, 5, o["price"]).number_format = kit.MONEY_FMT
        amount = o["qty"] * o["price"] if i < 2 else f"=D{r}*E{r}"
        ws.cell(r, 6, amount).number_format = kit.MONEY_FMT
    new_total = total_row + n
    ws.cell(new_total, 4).value = f"=SUM(D2:D{new_total - 1})"
    ws.cell(new_total, 6).value = f"=SUM(F2:F{new_total - 1})"
    wb.save(path)
