"""Reproduces the failure: rename the worksheet object and save."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "report.xlsx"
    wb = openpyxl.load_workbook(path)
    wb["Data"].title = "Sales 2024"
    wb.save(path)
