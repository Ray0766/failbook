"""A wrong solution: real dates and real numbers, the amounts formatted '#,##0.00' like the existing
rows, but the dates given a negative section ('yyyy-mm-dd;-yyyy-mm-dd'). The new March dates look the
same either way, but the format is not the existing rows' yyyy-mm-dd: in LibreOffice it shows -1766.34
as -1904-10-31 where yyyy-mm-dd shows 1895-02-27.
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
        ws.cell(r, 1, dt.date.fromisoformat(row["date"])).number_format = "yyyy-mm-dd;-yyyy-mm-dd"
        ws.cell(r, 2, row["category"])
        ws.cell(r, 3, row["description"])
        ws.cell(r, 4, float(row["amount"].replace(",", ""))).number_format = "#,##0.00"
    wb.save(path)
