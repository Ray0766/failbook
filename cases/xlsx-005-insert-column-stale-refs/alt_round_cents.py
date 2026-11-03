"""A valid alternative: every Line Total rounded to whole cents (=ROUND(B*D*(1-C),2)), which is
normal for money amounts; the task does not ask for unrounded values. The Summary keeps summing
the (rounded) Line Totals.
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
        ws.cell(r, 3, discounts.get(line["product"], 0)).number_format = "0%"
        ws.cell(r, 5, f"=ROUND(B{r}*D{r}*(1-C{r}),2)")
    summary = wb["Summary"]
    summary["B1"] = f"=SUM(Lines!E2:E{last})"
    summary["B2"] = f"=SUM(Lines!D2:D{last})"
    summary["B3"] = f"=AVERAGE(Lines!B2:B{last})"
    wb.save(path)
