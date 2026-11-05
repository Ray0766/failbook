"""A correct solution: an absolute reference to the rate cell."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "invoice.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Invoice"]
    for r in range(4, 4 + len(meta["lines"])):
        ws.cell(r, 5, f"=D{r}*$B$1")
        ws.cell(r, 6, f"=D{r}+E{r}")
        ws.cell(r, 5).number_format = ws.cell(r, 6).number_format = "#,##0.00"
    wb.save(path)
