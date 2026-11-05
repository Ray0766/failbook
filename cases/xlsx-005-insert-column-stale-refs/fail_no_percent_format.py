"""A wrong solution: everything is correct and live, but the discount cells are left in the General
format, although the task asks for the discounts to be formatted as a percentage.
"""

import openpyxl


def solve(workdir, meta):
    path = workdir / "orders.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Lines"]
    ws.insert_cols(3)
    ws.cell(1, 3, "Discount")
    ws.cell(1, 3)._style = ws.cell(1, 2)._style
    discounts = {d["product"]: d["discount"] for d in meta["discounts"]}
    last = 1 + len(meta["lines"])
    for r, line in enumerate(meta["lines"], start=2):
        ws.cell(r, 3, discounts.get(line["product"], 0)).number_format = "General"
        ws.cell(r, 5, f"=B{r}*D{r}*(1-C{r})")
    summary = wb["Summary"]
    summary["B1"] = f"=SUM(Lines!E2:E{last})"
    summary["B2"] = f"=SUM(Lines!D2:D{last})"
    summary["B3"] = f"=AVERAGE(Lines!B2:B{last})"
    wb.save(path)
