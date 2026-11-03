"""A wrong solution: the new orders are inserted above the Total row with the right values, Amount
formulas and extended totals, but no number format is set on them. openpyxl stores the dates as
datetimes with its default 'yyyy-mm-dd h:mm:ss', and Unit Price and Amount stay General (12.5 and 87
where the other order rows show 12.50 and 87.00), so the rows are not "in the same format as the
other order rows".
"""

import datetime as dt

import openpyxl


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
        ws.cell(r, 2, dt.datetime.fromisoformat(o["date"]))
        ws.cell(r, 3, o["customer"])
        ws.cell(r, 4, o["qty"])
        ws.cell(r, 5, o["price"])
        ws.cell(r, 6, f"=D{r}*E{r}")
    new_total = total_row + n
    ws.cell(new_total, 4).value = f"=SUM(D2:D{new_total - 1})"
    ws.cell(new_total, 6).value = f"=SUM(F2:F{new_total - 1})"
    wb.save(path)
