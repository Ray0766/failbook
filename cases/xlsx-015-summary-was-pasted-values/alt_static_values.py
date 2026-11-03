"""Alternative valid solution: keep the Summary as plain values, but write the corrected numbers.

The task asks for correct totals, not for formulas, so this must pass.
"""

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
    totals = {}
    for r in range(2, ws.max_row + 1):
        totals[ws.cell(r, 1).value] = totals.get(ws.cell(r, 1).value, 0) + ws.cell(r, 3).value
    s = wb["Summary"]
    for r in range(2, s.max_row + 1):
        label = s.cell(r, 1).value
        if label == "Total":
            s.cell(r, 2).value = sum(totals.values())
        elif label in totals:
            s.cell(r, 2).value = totals[label]
    wb.save(path)
