"""A correct solution: rename the sheet and rewrite every reference to it (formulas and defined names)."""

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
    wb.save(path)
