"""A wrong solution: the growth targets are entered correctly as fractions, but the 2025 Target
column, the Total row and the overall growth figure are replaced by typed numbers. Today's figures
are right, but they no longer follow the revenue and growth targets, which the task asks for.
"""

import openpyxl


def solve(workdir, meta):
    path = workdir / "targets.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Targets"]
    total_rev = total_target = 0
    for r, t in enumerate(meta["targets"], start=2):
        ws.cell(r, 3).value = t["pct"] / 100
        target = t["revenue"] * (1 + t["pct"] / 100)
        ws.cell(r, 4).value = target
        total_rev += t["revenue"]
        total_target += target
    total = 2 + len(meta["targets"])
    ws.cell(total, 2).value = total_rev
    ws.cell(total, 4).value = total_target
    ws.cell(total, 3).value = total_target / total_rev - 1
    wb.save(path)
