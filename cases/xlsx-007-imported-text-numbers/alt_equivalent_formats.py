"""A valid alternative: real dates and real numbers, formatted in ways that differ from the existing
rows' formats only where no value these columns hold shows differently: the dates as 'yyyy-mm-dd;@' (a
text section that shows text as typed, as a format without one does) and the amounts as
'#,##0.00;-#,##0.00' (a negative section that shows every negative amount of a cent or more as
'#,##0.00' does; the amounts are positive). In LibreOffice every cell shows the same as with the
existing rows' formats.
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
        ws.cell(r, 1, dt.date.fromisoformat(row["date"])).number_format = "yyyy-mm-dd;@"
        ws.cell(r, 2, row["category"])
        ws.cell(r, 3, row["description"])
        ws.cell(r, 4, float(row["amount"].replace(",", ""))).number_format = "#,##0.00;-#,##0.00"
    wb.save(path)
