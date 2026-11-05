"""Wrong: Q2 and Year are formulas, but the Q1, Q3 and Q4 Grand Totals are typed in as fixed numbers."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "regional.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Sales"]
    last = ws.max_row
    total = last + 1
    ws.cell(total, 1, "Grand Total")
    reps = [rep for group in meta["groups"] for rep in group["reps"]]
    for i, col in enumerate("CDEF"):
        if col == "D":
            ws[f"{col}{total}"] = f'=SUMIF($B$2:$B${last},"<>",{col}2:{col}{last})'
        else:
            ws[f"{col}{total}"] = sum(rep["q"][i] for rep in reps)
    ws[f"G{total}"] = f'=SUMIF($B$2:$B${last},"<>",G2:G{last})'
    wb.save(path)
