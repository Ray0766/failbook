"""Alternative valid solution: move the Total row down by hand and use SUBTOTAL()."""

from copy import copy

import openpyxl

from failbook import kit


def solve(workdir, meta):
    path = workdir / "sales.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Orders"]
    total_row = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "Total")
    n = len(meta["new"])
    new_total = total_row + n
    for c in range(1, 7):
        src, dst = ws.cell(total_row, c), ws.cell(new_total, c)
        dst._style = src._style
        src.value = None
    for i, o in enumerate(meta["new"]):
        r = total_row + i
        ws.cell(r, 1, o["id"])
        ws.cell(r, 2, kit.from_iso(o["date"])).number_format = kit.DATE_FMT
        ws.cell(r, 3, o["customer"])
        ws.cell(r, 4, o["qty"])
        ws.cell(r, 5, o["price"]).number_format = kit.MONEY_FMT
        ws.cell(r, 6, f"=E{r}*D{r}").number_format = kit.MONEY_FMT
        ws.cell(r, 1).font = copy(ws.cell(2, 1).font)
    ws.cell(new_total, 1, "Total")
    ws.cell(new_total, 4, f"=SUBTOTAL(9,D2:D{new_total - 1})")
    ws.cell(new_total, 6, f"=SUBTOTAL(9,F2:F{new_total - 1})")
    wb.save(path)
