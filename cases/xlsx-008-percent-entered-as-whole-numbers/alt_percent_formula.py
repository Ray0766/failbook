"""A valid alternative: enter each growth target the way typing "8.5%" into a cell stores it when
written as a formula (=8.5%), and show two decimals (0.00%) instead of one.
"""

import openpyxl


def solve(workdir, meta):
    path = workdir / "targets.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Targets"]
    for r, t in enumerate(meta["targets"], start=2):
        ws.cell(r, 3).value = f"={t['pct']:g}%"
        ws.cell(r, 3).number_format = "0.00%"
    wb.save(path)
