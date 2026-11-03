"""A wrong solution: rename the sheet and rewrite every cell formula, but leave the defined name
RevenueRange pointing at the sheet that no longer exists (Data!). The Dashboard total that uses
the name evaluates to #NAME?, which the checker must grade as a failure, not as an engine gap.
"""

import re

import openpyxl

OLD = re.compile(r"(?<![A-Za-z0-9_.'])(?:'Data'|Data)!")
NEW = "'Sales 2024'!"


def solve(workdir, meta):
    path = workdir / "report.xlsx"
    wb = openpyxl.load_workbook(path)
    wb["Data"].title = "Sales 2024"
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.value = OLD.sub(NEW, cell.value)
    wb.save(path)
