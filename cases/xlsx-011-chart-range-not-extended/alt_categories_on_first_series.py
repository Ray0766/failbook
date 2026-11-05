"""Alternative valid solution: rebuild the chart over January to December, setting the month labels after
adding Revenue and before adding Costs, so only the Revenue series carries them. A line chart has one category
axis, drawn from the first series that has labels (LibreOffice renders Jan to Dec under both lines; its XLSX
import code says Excel 2007 and later do the same)."""

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
    chart.y_axis.title = "Amount"
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=last), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last))
    chart.add_data(Reference(ws, min_col=3, min_row=1, max_row=last), titles_from_data=True)
    chart.width, chart.height = 18, 9
    ws.add_chart(chart, "E2")
    wb.save(path)
