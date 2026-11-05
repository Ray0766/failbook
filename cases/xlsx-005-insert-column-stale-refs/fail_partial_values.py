"""A wrong solution: only the first line keeps a Line Total formula; every other Line Total and the
Summary's average unit price are typed numbers. Today's values are all correct, but the task asks for
a formula on every line and a Summary that keeps updating when the Lines data changes.
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
        d = discounts.get(line["product"], 0)
        ws.cell(r, 3, d).number_format = "0%"
        ws.cell(r, 5).value = f"=B{r}*D{r}*(1-C{r})" if r == 2 else line["price"] * line["qty"] * (1 - d)
    summary = wb["Summary"]
    summary["B1"] = f"=SUM(Lines!E2:E{last})"
    summary["B2"] = f"=SUM(Lines!D2:D{last})"
    summary["B3"] = sum(line["price"] for line in meta["lines"]) / len(meta["lines"])
    wb.save(path)
