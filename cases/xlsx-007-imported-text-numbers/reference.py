"""A correct solution: parse the dates and amounts before writing them."""

import csv
import datetime as dt
import io

import openpyxl


def solve(workdir, meta):
    path = workdir / "expenses.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Expenses"]
    reader = csv.DictReader(io.StringIO(meta["csv"]))
    r = ws.max_row
    for row in reader:
        r += 1
        ws.cell(r, 1, dt.date.fromisoformat(row["date"])).number_format = "yyyy-mm-dd"
        ws.cell(r, 2, row["category"])
        ws.cell(r, 3, row["description"])
        ws.cell(r, 4, float(row["amount"].replace(",", ""))).number_format = "#,##0.00"
    wb.save(path)
