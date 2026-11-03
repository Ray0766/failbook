"""Alternative valid solution: an algebraically equivalent formula, written as 1 - Cost/Price."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "pricing.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Products"]
    ws["E1"] = "Margin %"
    for r in range(2, 2 + len(meta["products"])):
        ws.cell(r, 5, f"=1-C{r}/D{r}").number_format = "0.00%"
    wb.save(path)
