"""A wrong solution: VAT and Gross are correct $B$1 formulas on every line and Total VAT and Total
Gross are SUM formulas, but Total Net (=SUM(D4:D15) in the file as given) is replaced by its value
typed in as a number. Every value shown is right and the rate still drives VAT and Gross, but Total
Net no longer follows the lines, which the task did not ask to change.
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
    ws.cell(total, 4).value = round(sum(line["qty"] * line["price"] for line in meta["lines"]), 2)
    ws.cell(total, 5).value = f"=SUM(E{first}:E{last})"
    ws.cell(total, 6).value = f"=SUM(F{first}:F{last})"
    wb.save(path)
