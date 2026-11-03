"""A correct solution: exact-match lookups (INDEX/MATCH here; VLOOKUP(...,FALSE) works too)."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "orders.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Orders"]
    last = 1 + len(meta["products"])
    keys = f"Products!$A$2:$A${last}"
    for r in range(2, 2 + len(meta["orders"])):
        ws.cell(r, 4, f"=INDEX(Products!$B$2:$B${last},MATCH(B{r},{keys},0))")
        ws.cell(r, 5, f"=VLOOKUP(B{r},Products!$A$2:$C${last},3,FALSE)").number_format = "#,##0.00"
        ws.cell(r, 6, f"=C{r}*E{r}").number_format = "#,##0.00"
    wb.save(path)
