# Evidence chain demo: xlsx-001

Output of `scripts/demo-evidence-chain.sh xlsx-001`, captured on 2026-11-03 on Ubuntu 24.04 with LibreOffice 24.2.7.2 and failbook 0.1.0. Temporary paths are shortened to `$WORK`. No AI model is involved: the failure is reproduced by `failure.py` and the corrected output by `reference.py`.

```text
$ failbook prepare xlsx-001 $WORK/attempt
Prepared xlsx-001 (seed 0) in $WORK/attempt
Give the agent the task in $WORK/attempt/TASK.md, then run:
  failbook grade xlsx-001 $WORK/attempt

$ failbook solve xlsx-001 $WORK/attempt --kind failure
Applied the failure solution of xlsx-001 to $WORK/attempt

$ failbook grade xlsx-001 $WORK/attempt
xlsx-001  New orders inserted above the Total row are left out of the total
verdict: FAIL   (seed 0, engine: LibreOffice 24.2.7.2 420(Build:2))
  [PASS] new-orders-added: Each new order appears once, with the given date, customer, qty and price
  [PASS] existing-orders-unchanged: The original orders are unchanged and in their original order
  [PASS] amounts-correct: Amount equals Qty x Unit Price on every order row
  [PASS] amounts-are-live: Every order's Amount follows its Qty and Unit Price (new rows and original rows)
  [FAIL] totals-include-all-orders: The Total row sums Qty and Amount over all orders, old and new
         Total Qty is 198, expected 269 (difference -71)
  [FAIL] totals-are-live: Changing the Qty and Unit Price of the orders updates Total Qty and Total Amount
         in a copy of the file where each order's Qty was raised by a different amount (+91 in all) and every Unit Price by 1.00, Total Qty changed by +55, expected +91 (difference -36)
  [PASS] total-row-last: The Total row is directly below the last order

$ failbook audit $WORK/attempt/sales.xlsx --before $WORK/original/sales.xlsx
[error] range-omits-adjacent-cells at Orders!D15: D2:D11 stops at row 11, but rows 12-14 between it and the formula hold values
[error] range-omits-adjacent-cells at Orders!F15: F2:F11 stops at row 11, but rows 12-14 between it and the formula hold values
[error] formula-moved-as-text at Orders!D15: 2 formula(s) were moved without updating their references, e.g. =SUM(D2:D11) was at D12

$ failbook solve xlsx-001 $WORK/fixed --kind reference
Applied the reference solution of xlsx-001 to $WORK/fixed

$ failbook grade xlsx-001 $WORK/fixed
xlsx-001  New orders inserted above the Total row are left out of the total
verdict: PASS   (seed 0, engine: LibreOffice 24.2.7.2 420(Build:2))
  [PASS] new-orders-added: Each new order appears once, with the given date, customer, qty and price
  [PASS] existing-orders-unchanged: The original orders are unchanged and in their original order
  [PASS] amounts-correct: Amount equals Qty x Unit Price on every order row
  [PASS] amounts-are-live: Every order's Amount follows its Qty and Unit Price (new rows and original rows)
  [PASS] totals-include-all-orders: The Total row sums Qty and Amount over all orders, old and new
  [PASS] totals-are-live: Changing the Qty and Unit Price of the orders updates Total Qty and Total Amount
  [PASS] total-row-last: The Total row is directly below the last order

$ failbook audit $WORK/fixed/sales.xlsx --before $WORK/original/sales.xlsx
no findings
```
