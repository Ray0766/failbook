"""A wrong solution: every line's VAT (=D4*$B$1) and Gross (=D4+E4) are right, but the Total row's VAT
and Gross (=SUM(E4:E15) and =SUM(F4:F15) in the file as given, for seed 0) are replaced by formulas on a
typed copy of the Net total (=21655.63*$B$1 and =21655.63+E16 for seed 0). Today's totals are right and
follow B1, but change a Qty or a Unit Price and Total Net moves while Total VAT and Total Gross keep the
old Net total.
"""

import openpyxl


def solve(workdir, meta):
    path = workdir / "invoice.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Invoice"]
    first, last = 4, 3 + len(meta["lines"])
    for r in range(first, last + 1):
        ws.cell(r, 5, f"=D{r}*$B$1").number_format = "#,##0.00"
        ws.cell(r, 6, f"=D{r}+E{r}").number_format = "#,##0.00"
    total = last + 1
    net_total = round(sum(line["qty"] * line["price"] for line in meta["lines"]), 2)
    ws.cell(total, 5).value = f"={net_total}*$B$1"
    ws.cell(total, 6).value = f"={net_total}+E{total}"
    wb.save(path)
