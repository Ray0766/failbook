"""A wrong solution: the Net formulas are kept, but each line's VAT formula multiplies a typed copy of
its Net by $B$1 (=95.47*$B$1) and Gross adds the same typed Net to VAT (=95.47+E4). Today's values are
right and changing B1 updates VAT and Gross, but they no longer follow the Net: change a Qty or a Unit
Price and Net moves while VAT and Gross keep the old Net, so VAT is no longer "Net x the VAT rate in
B1" and Gross no longer "Net + VAT".
"""

import openpyxl


def solve(workdir, meta):
    path = workdir / "invoice.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Invoice"]
    for r, line in enumerate(meta["lines"], start=4):
        net = round(line["qty"] * line["price"], 2)
        ws.cell(r, 5, f"={net}*$B$1").number_format = "#,##0.00"
        ws.cell(r, 6, f"={net}+E{r}").number_format = "#,##0.00"
    wb.save(path)
