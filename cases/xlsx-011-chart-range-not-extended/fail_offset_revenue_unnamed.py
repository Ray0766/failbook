"""Wrong: Costs and the month labels are extended, and Revenue is switched to a dynamic OFFSET name (which the
checker cannot resolve), but the Revenue series loses its name, so the legend no longer says which line is
Revenue. The checker cannot read what the OFFSET series plots, but it is the only series that can be the Revenue
line, and it has no name."""

import re

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
    for series in ws._charts[0].series:
        series.val.numRef.f = re.sub(r"\$(\d+)$", f"${last}", series.val.numRef.f)
        ref = series.cat.numRef or series.cat.strRef
        ref.f = re.sub(r"\$(\d+)$", f"${last}", ref.f)
    wb.defined_names["RevenueDyn"] = DefinedName(
        "RevenueDyn", attr_text="OFFSET(Monthly!$B$2,0,0,COUNTA(Monthly!$A:$A)-1,1)")
    revenue = ws._charts[0].series[0]
    revenue.val.numRef.f = "[0]!RevenueDyn"
    revenue.tx = None
    wb.save(path)
