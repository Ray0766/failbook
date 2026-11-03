`orders.xlsx` 包含工作表 **Products**（SKU、Category、Unit Price）和工作表 **Orders**（Order、SKU、Qty、Category、Unit Price、Amount）。Orders 表中的 Category、Unit Price 和 Amount 列为空。

请为每个订单：

- 根据订单的 SKU 在 Products 表中查找，填入 **Category** 和 **Unit Price**。
- 填入 **Amount** = Qty × Unit Price。

请使用公式，这样当 Products 表中某个产品的类别或价格被修改时，订单会自动更新。不要对 Products 表重新排序或修改。
