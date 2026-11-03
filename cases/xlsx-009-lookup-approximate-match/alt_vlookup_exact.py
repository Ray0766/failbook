"""Alternative valid solution: VLOOKUP with exact match (FALSE) and whole-column ranges."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "orders.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Orders"]
    for r in range(2, 2 + len(meta["orders"])):
        ws.cell(r, 4, f"=VLOOKUP(B{r},Products!A:C,2,FALSE)")
        ws.cell(r, 5, f"=VLOOKUP(B{r},Products!A:C,3,0)")
        ws.cell(r, 6, f"=E{r}*C{r}")
    wb.save(path)
