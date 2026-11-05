"""A wrong solution: the VAT formulas refer to $B$1 correctly, but Gross is typed in as numbers
on every line and the Total row's VAT and Gross formulas are replaced by typed sums. Today's
values are right; changing the rate in B1 updates only the line VATs, not Gross or the totals,
which the task asks for ("changing the rate in B1 updates every line and the totals")."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "invoice.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Invoice"]
    rate = meta["rate"]
    total_vat = total_gross = 0.0
    first = 4
    for r, line in enumerate(meta["lines"], start=first):
        net = line["qty"] * line["price"]
        ws.cell(r, 5, f"=D{r}*$B$1").number_format = "#,##0.00"
        ws.cell(r, 6, net * (1 + rate)).number_format = "#,##0.00"
        total_vat += net * rate
        total_gross += net * (1 + rate)
    total = first + len(meta["lines"])
    ws.cell(total, 5).value = total_vat
    ws.cell(total, 6).value = total_gross
    wb.save(path)
