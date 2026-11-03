"""Reproduces the failure: append the rows, leave validation and conditional formatting ranges alone."""

import openpyxl

from failbook import kit


def solve(workdir, meta):
    path = workdir / "tracker.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Tracker"]
    for t in meta["new"]:
        ws.append([t["id"], t["task"], t["owner"], t["status"], kit.from_iso(t["due"])])
        ws.cell(ws.max_row, 5).number_format = kit.DATE_FMT
    wb.save(path)
