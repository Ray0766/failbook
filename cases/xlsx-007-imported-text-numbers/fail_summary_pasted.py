"""A wrong solution: the new expenses are appended correctly (real dates and numbers), but every
Summary formula is replaced by its current total as a typed number. Today's totals are right, but
the Summary no longer follows the expense log, and the task did not ask to change it.
"""

import csv
import datetime as dt
import io

import openpyxl

CATEGORIES = ["Travel", "Meals", "Office", "Software", "Training"]
MONTHS = {"January": 1, "February": 2, "March": 3}


def solve(workdir, meta):
    path = workdir / "expenses.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Expenses"]
    r = 1 + len(meta["existing"])
    for row in csv.DictReader(io.StringIO(meta["csv"])):
        r += 1
        ws.cell(r, 1, dt.date.fromisoformat(row["date"])).number_format = "yyyy-mm-dd"
        ws.cell(r, 2, row["category"])
        ws.cell(r, 3, row["description"])
        ws.cell(r, 4, float(row["amount"].replace(",", ""))).number_format = "#,##0.00"
    every = meta["existing"] + meta["new"]
    summary = wb["Summary"]
    for rr in range(1, summary.max_row + 1):
        label = summary.cell(rr, 1).value
        if label in CATEGORIES:
            summary.cell(rr, 2).value = sum(e["amount"] for e in every if e["category"] == label)
        elif label in MONTHS:
            summary.cell(rr, 2).value = sum(e["amount"] for e in every if int(e["date"][5:7]) == MONTHS[label])
        elif label == "Grand total":
            summary.cell(rr, 2).value = sum(e["amount"] for e in every)
    wb.save(path)
