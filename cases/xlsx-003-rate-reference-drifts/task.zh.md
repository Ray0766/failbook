`invoice.xlsx` 是一张发票草稿（工作表 **Invoice**），增值税率在 **B1** 单元格。每个明细行已经有 **Net**（不含税金额），**VAT** 和 **Gross** 两列为空。

请为每个明细行填写：

- **VAT** = Net × B1 中的税率
- **Gross** = Net + VAT

请使用公式，这样修改 B1 的税率时，每一行和合计都会随之更新。底部的 **Total** 行必须显示正确的 Net、VAT、Gross 合计。
