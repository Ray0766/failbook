"""A wrong solution: live formulas, but each margin is rounded to four decimals with ROUND().
The task gives the exact formula and says not to round the margin (the percentage format decides
how many decimals are shown), so the checker must grade this fail, with messages that show the
small difference rather than two numbers that look the same."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "pricing.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Products"]
    ws["E1"] = "Margin %"
    for r in range(2, 2 + len(meta["products"])):
        ws.cell(r, 5, f"=ROUND((D{r}-C{r})/D{r},4)").number_format = "0.00%"
    wb.save(path)
