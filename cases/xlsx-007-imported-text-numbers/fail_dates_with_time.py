"""A wrong solution: real dates and real numbers, the amounts formatted '#,##0.00' like the existing
rows, but the dates given a format with a time part ('yyyy-mm-dd hh:mm', shown as 2024-03-02 00:00),
so the dates are not "formatted like the existing rows" (yyyy-mm-dd).
"""

import csv
import datetime as dt
import io

import openpyxl


def solve(workdir, meta):
    path = workdir / "expenses.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Expenses"]
    r = 1 + len(meta["existing"])
    for row in csv.DictReader(io.StringIO(meta["csv"])):
        r += 1
        ws.cell(r, 1, dt.date.fromisoformat(row["date"])).number_format = "yyyy-mm-dd hh:mm"
        ws.cell(r, 2, row["category"])
        ws.cell(r, 3, row["description"])
        ws.cell(r, 4, float(row["amount"].replace(",", ""))).number_format = "#,##0.00"
    wb.save(path)
