"""Wrong: Revenue is switched to a dynamic OFFSET name (which the checker cannot resolve), but the Costs series
and its month labels still stop at October. The unreadable Revenue series must not hide the stale Costs series."""

import openpyxl
from openpyxl.workbook.defined_name import DefinedName


def solve(workdir, meta):
    path = workdir / "monthly.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Monthly"]
    for m in meta["new"]:
        ws.append([m["month"], m["revenue"], m["costs"]])
        ws.cell(ws.max_row, 2).number_format = ws.cell(ws.max_row, 3).number_format = "#,##0"
    for name, col in (("MonthsDyn", "A"), ("RevenueDyn", "B")):
        wb.defined_names[name] = DefinedName(name, attr_text=f"OFFSET(Monthly!${col}$2,0,0,COUNTA(Monthly!$A:$A)-1,1)")
    revenue = ws._charts[0].series[0]
    revenue.val.numRef.f = "[0]!RevenueDyn"
    (revenue.cat.numRef or revenue.cat.strRef).f = "[0]!MonthsDyn"
    wb.save(path)
