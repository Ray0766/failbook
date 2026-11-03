"""Wrong: Category is typed in as fixed text; only Unit Price is a lookup, so a category change on Products is lost."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "orders.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Orders"]
    category = {p["sku"]: p["category"] for p in meta["products"]}
    for r, order in enumerate(meta["orders"], start=2):
        ws.cell(r, 4, category[order["sku"]])
        ws.cell(r, 5, f"=VLOOKUP(B{r},Products!$A:$C,3,FALSE)")
        ws.cell(r, 6, f"=C{r}*E{r}")
    wb.save(path)
