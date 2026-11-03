"""Alternative valid solution: point the existing chart's series at defined names that cover January to December."""

import openpyxl
from openpyxl.workbook.defined_name import DefinedName


def solve(workdir, meta):
    path = workdir / "monthly.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Monthly"]
    for m in meta["new"]:
        ws.append([m["month"], m["revenue"], m["costs"]])
        ws.cell(ws.max_row, 2).number_format = ws.cell(ws.max_row, 3).number_format = "#,##0"
    last = ws.max_row
    for name, col in (("ChartMonths", "A"), ("ChartRevenue", "B"), ("ChartCosts", "C")):
        wb.defined_names[name] = DefinedName(name, attr_text=f"Monthly!${col}$2:${col}${last}")
    chart = ws._charts[0]
    for series, name in zip(chart.series, ("ChartRevenue", "ChartCosts")):
        series.val.numRef.f = f"[0]!{name}"
        (series.cat.numRef or series.cat.strRef).f = "[0]!ChartMonths"
    wb.save(path)
