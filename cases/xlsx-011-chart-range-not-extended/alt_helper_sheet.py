"""Alternative valid solution: the chart reads its values and month labels from a helper sheet ``ChartData``
whose cells are formulas pointing at the Monthly sheet (=Monthly!B2 and so on), extended to December. The series
are not over the Monthly columns, but the cells they plot hold the January to December Revenue and Costs figures
and the labels read Jan to Dec."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "monthly.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Monthly"]
    for m in meta["new"]:
        ws.append([m["month"], m["revenue"], m["costs"]])
        ws.cell(ws.max_row, 2).number_format = ws.cell(ws.max_row, 3).number_format = "#,##0"
    last = ws.max_row
    helper = wb.create_sheet("ChartData")
    for r in range(1, last + 1):
        for col in "ABC":
            helper[f"{col}{r}"] = f"=Monthly!{col}{r}"
    revenue, costs = ws._charts[0].series
    for series, col in ((revenue, "B"), (costs, "C")):
        series.val.numRef.f = f"'ChartData'!${col}$2:${col}${last}"
        (series.cat.numRef or series.cat.strRef).f = f"'ChartData'!$A$2:$A${last}"
    wb.save(path)
