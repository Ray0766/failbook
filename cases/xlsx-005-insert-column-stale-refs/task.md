`orders.xlsx` has a sheet **Lines** (Product, Unit Price, Qty, Line Total) and a sheet **Summary** that reports totals for the Lines sheet.

Insert a new column **Discount** between **Unit Price** and **Qty**, so that the Lines columns become: Product, Unit Price, Discount, Qty, Line Total.

Enter these discounts (as fractions, e.g. 5% is 0.05, formatted as a percentage). Products that are not listed get a discount of 0:

{{discounts_md}}

Then:

- Change **Line Total** to Unit Price × Qty × (1 − Discount), as a formula on every line.
- The **Summary** sheet must keep showing the correct total revenue (after discounts), total units and average unit price, and keep updating when the Lines data changes.
