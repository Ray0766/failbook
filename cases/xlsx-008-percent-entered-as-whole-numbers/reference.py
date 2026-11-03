"""A correct solution: a percentage is a fraction."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "targets.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Targets"]
    for r, t in enumerate(meta["targets"], start=2):
        ws.cell(r, 3).value = t["pct"] / 100
    wb.save(path)
