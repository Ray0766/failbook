"""Alternative valid solution: VAT rounded to whole cents on each line (common invoice practice),
Gross = Net + the rounded VAT, and the existing Total row summing the rounded lines. The task does
not ask for unrounded amounts, so this must pass."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "invoice.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Invoice"]
    for r in range(4, 4 + len(meta["lines"])):
        ws.cell(r, 5, f"=ROUND(D{r}*$B$1,2)")
        ws.cell(r, 6, f"=D{r}+E{r}")
        ws.cell(r, 5).number_format = ws.cell(r, 6).number_format = "#,##0.00"
    wb.save(path)
