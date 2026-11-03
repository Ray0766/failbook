"""Reproduces the failure: VLOOKUP without the range_lookup argument (approximate match)."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "orders.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Orders"]
    last = 1 + len(meta["products"])
    for r in range(2, 2 + len(meta["orders"])):
        ws.cell(r, 4, f"=VLOOKUP(B{r},Products!$A$2:$C${last},2)")
        ws.cell(r, 5, f"=VLOOKUP(B{r},Products!$A$2:$C${last},3)").number_format = "#,##0.00"
        ws.cell(r, 6, f"=C{r}*E{r}").number_format = "#,##0.00"
    wb.save(path)
