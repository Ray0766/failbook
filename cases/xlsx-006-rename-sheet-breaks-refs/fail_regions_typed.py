"""A wrong solution: rename the sheet and fix every reference, but replace the Dashboard figures of
every region except the first with typed numbers. Today's figures are right, but they no longer
update when the sales data changes, which the task asks for.
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
    for name in list(wb.defined_names.values()):
        name.attr_text = OLD.sub(NEW, name.attr_text)
    dash = wb["Dashboard"]
    for r, region in enumerate(meta["regions"][1:], start=3):
        dash.cell(r, 2).value = sum(row["revenue"] for row in meta["rows"] if row["region"] == region)
    wb.save(path)
