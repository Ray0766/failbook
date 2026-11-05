"""Alternative valid solution: give B1 a name and refer to it by name."""

import openpyxl
from openpyxl.workbook.defined_name import DefinedName


def solve(workdir, meta):
    path = workdir / "invoice.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Invoice"]
    wb.defined_names["VatRate"] = DefinedName("VatRate", attr_text="Invoice!$B$1")
    for r in range(4, 4 + len(meta["lines"])):
        ws.cell(r, 5, f"=ROUND(D{r},10)*VatRate")
        ws.cell(r, 6, f"=D{r}*(1+VatRate)")
    wb.save(path)
