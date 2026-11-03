"""A wrong solution (the openpyxl slip of xlsx-001's fail_data_only_roundtrip.py): the workbook is
opened with data_only=True to read the Net amounts, so saving it turns every Net formula (=B4*C4 ...)
and the Total Net formula (=SUM(D4:D15)) into typed numbers. VAT and Gross get correct $B$1 formulas
and Total VAT and Total Gross are rewritten as SUM formulas, so every value shown is right and the rate
in B1 still drives VAT and Gross, but Net no longer follows Qty and Unit Price, which the task did not
ask to change.
"""

import openpyxl


def solve(workdir, meta):
    path = workdir / "invoice.xlsx"
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Invoice"]
    first, last = 4, 3 + len(meta["lines"])
    for r in range(first, last + 1):
        ws.cell(r, 5, f"=D{r}*$B$1").number_format = "#,##0.00"
        ws.cell(r, 6, f"=D{r}+E{r}").number_format = "#,##0.00"
    total = last + 1
    ws.cell(total, 5).value = f"=SUM(E{first}:E{last})"
    ws.cell(total, 6).value = f"=SUM(F{first}:F{last})"
    wb.save(path)
