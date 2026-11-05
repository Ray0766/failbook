"""Reproduces the failure: write the CSV fields into the sheet exactly as they appear (strings)."""

import csv
import io

import openpyxl


def solve(workdir, meta):
    path = workdir / "expenses.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Expenses"]
    reader = csv.reader(io.StringIO(meta["csv"]))
    next(reader)  # header
    for row in reader:
        ws.append(row)
    wb.save(path)
