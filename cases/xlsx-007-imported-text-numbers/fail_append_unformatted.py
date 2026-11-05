"""A wrong solution, the common openpyxl ws.append() approach: the dates are parsed into real dates
and the amounts into real numbers, but no number format is copied from the existing rows. The dates
show with openpyxl's default datetime format 'yyyy-mm-dd h:mm:ss' and the amounts in General (1766.34
instead of 1,766.34), so they are not "formatted like the existing rows".
"""

import csv
import datetime as dt
import io

import openpyxl


def solve(workdir, meta):
    path = workdir / "expenses.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Expenses"]
    for row in csv.DictReader(io.StringIO(meta["csv"])):
        ws.append([dt.datetime.fromisoformat(row["date"]), row["category"], row["description"],
                   float(row["amount"].replace(",", ""))])
    wb.save(path)
