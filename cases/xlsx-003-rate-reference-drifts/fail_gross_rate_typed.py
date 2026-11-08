"""A wrong solution: VAT refers to $B$1, but each Gross formula has the rate typed in
(=D4*(1+0.19)), so changing the rate in B1 leaves every Gross and the Gross total unchanged."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "invoice.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Invoice"]
    for r in range(4, 4 + len(meta["lines"])):
        ws.cell(r, 5, f"=D{r}*$B$1")
        ws.cell(r, 6, f"=D{r}*(1+{meta['rate']})")
        ws.cell(r, 5).number_format = ws.cell(r, 6).number_format = "#,##0.00"
    wb.save(path)
