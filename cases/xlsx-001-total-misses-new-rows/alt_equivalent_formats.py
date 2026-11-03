"""A valid alternative: the new rows get formats that differ from the other order rows' only in ways
that change no value these columns hold: the date as '[$-409]yyyy-mm-dd' (a locale tag on a date of
digits and '-', which it does not change), Unit Price as '#,##0.00;-#,##0.00' (a negative section that
shows every negative price of a cent or more as '#,##0.00' does; the prices are positive) and Amount as
'#,##0.00;@' (a text section that shows text as typed). In LibreOffice every cell shows the same as with
the formats of the other order rows.
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
        ws.cell(r, 2, kit.from_iso(o["date"])).number_format = "[$-409]yyyy-mm-dd"
        ws.cell(r, 3, o["customer"])
        ws.cell(r, 4, o["qty"])
        ws.cell(r, 5, o["price"]).number_format = "#,##0.00;-#,##0.00"
        ws.cell(r, 6, f"=D{r}*E{r}").number_format = "#,##0.00;@"
    new_total = total_row + n
    ws.cell(new_total, 4).value = f"=SUM(D2:D{new_total - 1})"
    ws.cell(new_total, 6).value = f"=SUM(F2:F{new_total - 1})"
    wb.save(path)
