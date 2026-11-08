"""A correct solution: sort the data and rewrite each Bonus formula for its new row."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "staff.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Staff"]
    rows = [[c.value for c in row][:4] for row in ws.iter_rows(min_row=2, max_row=ws.max_row)]
    rows.sort(key=lambda row: (row[1].casefold(), row[0].casefold()))
    for r, row in enumerate(rows, start=2):
        for c, value in enumerate(row, start=1):
            ws.cell(r, c).value = value
        ws.cell(r, 5).value = f"=C{r}*D{r}"
    wb.save(path)
