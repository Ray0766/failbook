"""A correct solution: one formula per row."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "pricing.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Products"]
    ws.cell(1, 5, "Margin %")
    ws.cell(1, 5)._style = ws.cell(1, 4)._style
    for r in range(2, 2 + len(meta["products"])):
        ws.cell(r, 5, f"=(D{r}-C{r})/D{r}").number_format = "0.0%"
    wb.save(path)
