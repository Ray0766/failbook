"""Reproduces the failure: insert rows with openpyxl and fill them in, leaving the totals alone.

openpyxl's insert_rows() shifts the cells below the insertion point but never
rewrites formulas, so the Total row keeps SUM(D2:D<old last row>).
"""

import openpyxl

from failbook import kit


def solve(workdir, meta):
    path = workdir / "sales.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Orders"]
    total_row = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "Total")
    ws.insert_rows(total_row, amount=len(meta["new"]))
    for i, o in enumerate(meta["new"]):
        r = total_row + i
        ws.cell(r, 1, o["id"])
        ws.cell(r, 2, kit.from_iso(o["date"])).number_format = kit.DATE_FMT
        ws.cell(r, 3, o["customer"])
        ws.cell(r, 4, o["qty"])
        ws.cell(r, 5, o["price"]).number_format = kit.MONEY_FMT
        ws.cell(r, 6, f"=D{r}*E{r}").number_format = kit.MONEY_FMT
    wb.save(path)
