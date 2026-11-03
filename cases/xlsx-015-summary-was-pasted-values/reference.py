"""A correct solution: apply the corrections and replace the pasted totals with live formulas."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "budget.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Q3 Data"]
    fixes = {(c["dept"], c["month"]): c["new"] for c in meta["corrections"]}
    last = ws.max_row
    for r in range(2, last + 1):
        key = (ws.cell(r, 1).value, ws.cell(r, 2).value)
        if key in fixes:
            ws.cell(r, 3).value = fixes[key]
    s = wb["Summary"]
    first = None
    for r in range(2, s.max_row + 1):
        label = s.cell(r, 1).value
        if label == "Total":
            s.cell(r, 2).value = f"=SUM(B{first}:B{r - 1})"
        elif label:
            first = first or r
            s.cell(r, 2).value = f"=SUMIF('Q3 Data'!$A$2:$A${last},A{r},'Q3 Data'!$C$2:$C${last})"
    wb.save(path)
