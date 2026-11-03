"""Wrong: the chart is moved onto a helper sheet ``ChartData`` that mirrors the Monthly sheet only up to October,
so both lines still end at October although the new rows are on the Monthly sheet."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "monthly.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Monthly"]
    stale = ws.max_row
    for m in meta["new"]:
        ws.append([m["month"], m["revenue"], m["costs"]])
        ws.cell(ws.max_row, 2).number_format = ws.cell(ws.max_row, 3).number_format = "#,##0"
    helper = wb.create_sheet("ChartData")
    for r in range(1, stale + 1):
        for col in "ABC":
            helper[f"{col}{r}"] = f"=Monthly!{col}{r}"
    for series, col in zip(ws._charts[0].series, "BC"):
        series.val.numRef.f = f"'ChartData'!${col}$2:${col}${stale}"
        (series.cat.numRef or series.cat.strRef).f = f"'ChartData'!$A$2:$A${stale}"
    wb.save(path)
