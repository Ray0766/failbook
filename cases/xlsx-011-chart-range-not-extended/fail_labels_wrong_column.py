"""Wrong: the series are extended, but the month labels now point at the Revenue column."""

import re

import openpyxl


def solve(workdir, meta):
    path = workdir / "monthly.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Monthly"]
    for m in meta["new"]:
        ws.append([m["month"], m["revenue"], m["costs"]])
        ws.cell(ws.max_row, 2).number_format = ws.cell(ws.max_row, 3).number_format = "#,##0"
    last = ws.max_row
    for chart in ws._charts:
        for series in chart.series:
            series.val.numRef.f = re.sub(r"\$(\d+)$", f"${last}", series.val.numRef.f)
            ref = series.cat.numRef or series.cat.strRef
            ref.f = f"Monthly!$B$2:$B${last}"
    wb.save(path)
