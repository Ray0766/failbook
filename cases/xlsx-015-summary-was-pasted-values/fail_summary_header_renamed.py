"""Wrong: the corrections and live Summary totals are right, but the Summary header 'Q3 total' is renamed."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "budget.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Q3 Data"]
    fixes = {(c["dept"], c["month"]): c["new"] for c in meta["corrections"]}
    for r in range(2, ws.max_row + 1):
        key = (ws.cell(r, 1).value, ws.cell(r, 2).value)
        if key in fixes:
            ws.cell(r, 3).value = fixes[key]
    last = ws.max_row
    s = wb["Summary"]
    for r in range(2, 2 + len(meta["depts"])):
        s.cell(r, 2).value = f"=SUMIF('Q3 Data'!$A$2:$A${last},A{r},'Q3 Data'!$C$2:$C${last})"
    t = 2 + len(meta["depts"])
    s.cell(t, 2).value = f"=SUM(B2:B{t - 1})"
    s["B1"] = "Q3 total (corrected)"
    wb.save(path)
