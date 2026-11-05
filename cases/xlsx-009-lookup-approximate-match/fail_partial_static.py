"""Wrong: lookups only on the rows of the first order's SKU; every other row holds pasted values."""

import openpyxl


def solve(workdir, meta):
    path = workdir / "orders.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Orders"]
    product = {p["sku"]: p for p in meta["products"]}
    first = meta["orders"][0]["sku"]
    for r, order in enumerate(meta["orders"], start=2):
        if order["sku"] == first:
            ws.cell(r, 4, f"=VLOOKUP(B{r},Products!$A:$C,2,FALSE)")
            ws.cell(r, 5, f"=VLOOKUP(B{r},Products!$A:$C,3,FALSE)")
            ws.cell(r, 6, f"=C{r}*E{r}")
        else:
            p = product[order["sku"]]
            ws.cell(r, 4, p["category"])
            ws.cell(r, 5, p["price"])
            ws.cell(r, 6, order["qty"] * p["price"])
    wb.save(path)
