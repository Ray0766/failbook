`orders.xlsx` has a sheet **Products** (SKU, Category, Unit Price) and a sheet **Orders** (Order, SKU, Qty, Category, Unit Price, Amount). In Orders, the Category, Unit Price and Amount columns are empty.

For every order:

- Fill in **Category** and **Unit Price** by looking up the order's SKU in the Products sheet.
- Fill in **Amount** = Qty × Unit Price.

Use formulas, so that the orders update automatically if a product's category or price is changed in the Products sheet. Do not reorder or edit the Products sheet.
