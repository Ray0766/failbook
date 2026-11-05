"""Reproduces the failure: write the first row's formulas, then fill them down with relative references."""

import openpyxl
from openpyxl.formula.translate import Translator


def solve(workdir, meta):
    path = workdir / "invoice.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Invoice"]
    first = 4
    last = first + len(meta["lines"]) - 1
    vat, gross = "=D4*B1", "=D4+E4"
    for r in range(first, last + 1):
        ws.cell(r, 5, Translator(vat, origin="E4").translate_formula(f"E{r}"))
        ws.cell(r, 6, Translator(gross, origin="F4").translate_formula(f"F{r}"))
        ws.cell(r, 5).number_format = ws.cell(r, 6).number_format = "#,##0.00"
    wb.save(path)
