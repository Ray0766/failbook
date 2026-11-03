"""Reproduces the failure: openpyxl insert_cols(), fix the edited sheet, forget the Summary sheet."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "orders.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Lines"]
    ws.insert_cols(3)
    ws.cell(1, 3, "Discount")
    ws.cell(1, 3)._style = ws.cell(1, 2)._style
    discounts = {d["product"]: d["discount"] for d in meta["discounts"]}
    for r, line in enumerate(meta["lines"], start=2):
        ws.cell(r, 3, discounts.get(line["product"], 0)).number_format = "0%"
        ws.cell(r, 5, f"=B{r}*D{r}*(1-C{r})")
    wb.save(path)
