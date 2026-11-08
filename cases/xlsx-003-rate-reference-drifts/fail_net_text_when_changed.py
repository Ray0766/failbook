"""A wrong solution: every VAT (=D4*$B$1) and Gross (=D4+E4) is right, but the third line's Net is
replaced by a formula that only works for today's Qty (=IF(B6=<today's Qty>,B6*C6,"n/a")). Change that
Qty and the Net shows n/a, so it no longer follows its Qty and Unit Price (net-still-calculated). Its VAT
and Gross then show #VALUE!; they carry the bad Net and are not reported a second time.
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
    r = first + 2
    ws.cell(r, 4).value = f'=IF(B{r}={meta["lines"][2]["qty"]},B{r}*C{r},"n/a")'
    wb.save(path)
