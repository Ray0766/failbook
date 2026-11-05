"""A valid alternative: the new rows get the formats of the other order rows written another way:
the date as 'YYYY"-"MM"-"DD' (upper case, quoted dash) where the file has 'yyyy\\-mm\\-dd', and the
money cells with the format string read from the row above. Each shows the same as the other rows.
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
    above = total_row - 1
    for i, o in enumerate(meta["new"]):
        r = total_row + i
        ws.cell(r, 1, o["id"])
        ws.cell(r, 2, kit.from_iso(o["date"])).number_format = 'YYYY"-"MM"-"DD'
        ws.cell(r, 3, o["customer"])
        ws.cell(r, 4, o["qty"])
        ws.cell(r, 5, o["price"]).number_format = ws.cell(above, 5).number_format
        ws.cell(r, 6, f"=D{r}*E{r}").number_format = ws.cell(above, 6).number_format
    new_total = total_row + n
    ws.cell(new_total, 4).value = f"=SUM(D2:D{new_total - 1})"
    ws.cell(new_total, 6).value = f"=SUM(F2:F{new_total - 1})"
    wb.save(path)
