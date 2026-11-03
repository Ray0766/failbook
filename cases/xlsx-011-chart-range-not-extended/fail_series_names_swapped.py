"""Wrong: the series are extended, but their names are swapped (the legend calls the Costs line Revenue)."""

import re

import openpyxl
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.series import SeriesLabel


def solve(workdir, meta):
    path = workdir / "monthly.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Monthly"]
    for m in meta["new"]:
        ws.append([m["month"], m["revenue"], m["costs"]])
        ws.cell(ws.max_row, 2).number_format = ws.cell(ws.max_row, 3).number_format = "#,##0"
    last = ws.max_row
    chart = ws._charts[0]
    for series in chart.series:
        series.val.numRef.f = re.sub(r"\$(\d+)$", f"${last}", series.val.numRef.f)
        ref = series.cat.numRef or series.cat.strRef
        ref.f = re.sub(r"\$(\d+)$", f"${last}", ref.f)
    chart.series[0].tx = SeriesLabel(strRef=StrRef("'Monthly'!$C$1"))
    chart.series[1].tx = SeriesLabel(strRef=StrRef("'Monthly'!$B$1"))
    wb.save(path)
