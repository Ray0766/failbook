"""Reproduces the failure: write "12%" as the number 12 into a percentage-formatted cell."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "targets.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Targets"]
    for r, t in enumerate(meta["targets"], start=2):
        ws.cell(r, 3).value = t["pct"]
    wb.save(path)
