"""A wrong solution (a common openpyxl slip): the workbook is opened with data_only=True, which
loads the stored results instead of the formulas, so saving it turns every original Amount formula
into a typed number. The new rows and the Total row get proper formulas, and every value shown is
right, but the original orders' Amounts no longer follow their Qty and Unit Price.
"""

import openpyxl

from failbook import kit


def solve(workdir, meta):
    path = workdir / "sales.xlsx"
    wb = openpyxl.load_workbook(path, data_only=True)
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
        ws.cell(r, 6, f"=D{r}*E{r}").number_format = kit.MONEY_FMT
    new_total = total_row + n
    ws.cell(new_total, 4).value = f"=SUM(D2:D{new_total - 1})"
    ws.cell(new_total, 6).value = f"=SUM(F2:F{new_total - 1})"
    wb.save(path)
