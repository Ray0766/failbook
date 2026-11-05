"""A correct solution: append the rows, then point the existing chart's series at the full range."""

import re

import openpyxl

ROWS = re.compile(r"\$(\d+)$")


def _extend(formula, last):
    # 'Monthly'!$B$2:$B$11 -> 'Monthly'!$B$2:$B$13
    return ROWS.sub(f"${last}", formula)


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
            series.val.numRef.f = _extend(series.val.numRef.f, last)
            if series.cat is not None:
                ref = series.cat.numRef or series.cat.strRef
                ref.f = _extend(ref.f, last)
    wb.save(path)
