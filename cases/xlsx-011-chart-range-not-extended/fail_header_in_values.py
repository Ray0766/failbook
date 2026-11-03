"""Wrong: the chart is rebuilt with the values starting at the header row (``add_data`` without
``titles_from_data``), so each series has 13 points against 12 month labels and no names."""

import openpyxl
from openpyxl.chart import LineChart, Reference


def solve(workdir, meta):
    path = workdir / "monthly.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Monthly"]
    for m in meta["new"]:
        ws.append([m["month"], m["revenue"], m["costs"]])
        ws.cell(ws.max_row, 2).number_format = ws.cell(ws.max_row, 3).number_format = "#,##0"
    last = ws.max_row
    ws._charts = []
    chart = LineChart()
    chart.title = meta["title"]
    chart.add_data(Reference(ws, min_col=2, max_col=3, min_row=1, max_row=last))
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last))
    ws.add_chart(chart, "E2")
    wb.save(path)
