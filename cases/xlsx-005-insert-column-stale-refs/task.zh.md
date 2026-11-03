`orders.xlsx` 包含工作表 **Lines**（Product、Unit Price、Qty、Line Total）和工作表 **Summary**（汇总 Lines 表的数据）。

请在 **Unit Price** 和 **Qty** 之间插入新列 **Discount**，使 Lines 表的列依次为：Product、Unit Price、Discount、Qty、Line Total。

填入以下折扣（以小数表示，例如 5% 填 0.05，并设为百分比格式）。未列出的产品折扣为 0：

{{discounts_md}}

然后：

- 将每一行的 **Line Total** 改为公式：Unit Price × Qty × (1 − Discount)。
- **Summary** 表必须继续显示正确的总收入（折后）、总数量和平均单价，并在 Lines 数据变化时继续自动更新。
