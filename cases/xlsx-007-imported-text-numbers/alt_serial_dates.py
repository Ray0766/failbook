"""A valid alternative: store each new date as its serial number with the date format of the
existing rows, parse the amounts with Decimal, and copy the cell styles of the first data row.
"""

import csv
import datetime as dt
import io
from decimal import Decimal

import openpyxl


def solve(workdir, meta):
    path = workdir / "expenses.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Expenses"]
    r = 1 + len(meta["existing"])
    for row in csv.DictReader(io.StringIO(meta["csv"])):
        r += 1
        ws.cell(r, 1).value = (dt.date.fromisoformat(row["date"]) - dt.date(1899, 12, 30)).days
        ws.cell(r, 2).value = row["category"]
        ws.cell(r, 3).value = row["description"]
        ws.cell(r, 4).value = float(Decimal(row["amount"].replace(",", "")))
        for c in range(1, 5):
            ws.cell(r, c)._style = ws.cell(2, c)._style
    wb.save(path)
