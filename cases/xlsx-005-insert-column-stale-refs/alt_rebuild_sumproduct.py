"""A valid alternative: move Qty and Line Total one column to the right by hand (no insert_cols),
compute Line Total as B*D - B*D*C, and rewrite the Summary with SUMPRODUCT over the inputs and
whole-column references.
"""

import openpyxl


def solve(workdir, meta):
    path = workdir / "orders.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Lines"]
    discounts = {d["product"]: d["discount"] for d in meta["discounts"]}
    last = 1 + len(meta["lines"])
    for r in range(1, last + 1):
        qty_style, total_style = ws.cell(r, 3)._style, ws.cell(r, 4)._style
        ws.cell(r, 4).value = ws.cell(r, 3).value
        ws.cell(r, 4)._style = qty_style
        ws.cell(r, 5)._style = total_style
    ws.cell(1, 3, "Discount")._style = ws.cell(1, 2)._style
    ws.cell(1, 5, "Line Total")
    for r, line in enumerate(meta["lines"], start=2):
        cell = ws.cell(r, 3, discounts.get(line["product"], 0))
        cell.number_format = "0.0%"
        ws.cell(r, 5, f"=B{r}*D{r}-B{r}*D{r}*C{r}")
    summary = wb["Summary"]
    summary["B1"] = f"=SUMPRODUCT(Lines!B2:B{last},Lines!D2:D{last},1-Lines!C2:C{last})"
    summary["B2"] = "=SUM(Lines!D:D)"
    summary["B3"] = "=AVERAGE(Lines!B:B)"
    wb.save(path)
