"""Alternative valid solution: the chart's series values, month labels and series names are rewritten with the
sheet name in lower case ('monthly'!$B$2:$B$13, 'monthly'!$A$2:$A$13, 'monthly'!$B$1). Excel and LibreOffice
match sheet names in references in any letter case, so this is the same chart as the reference solution's."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "monthly.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Monthly"]
    for m in meta["new"]:
        ws.append([m["month"], m["revenue"], m["costs"]])
        ws.cell(ws.max_row, 2).number_format = ws.cell(ws.max_row, 3).number_format = "#,##0"
    last = ws.max_row
    revenue, costs = ws._charts[0].series
    for series, col in ((revenue, "B"), (costs, "C")):
        series.val.numRef.f = f"'monthly'!${col}$2:${col}${last}"
        (series.cat.numRef or series.cat.strRef).f = f"'monthly'!$A$2:$A${last}"
        series.tx.strRef.f = f"'monthly'!${col}$1"
    wb.save(path)
