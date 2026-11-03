---
name: spreadsheet-verify
description: Use when creating or editing .xlsx/.xlsm spreadsheets with code (openpyxl, pandas, xlsxwriter or raw XML). Lists the edits that silently break formulas, charts, validation and formatting, and how to recalculate and verify the workbook before reporting that the work is done.
---

# Editing spreadsheets safely

A spreadsheet edit is not done when the code runs without errors. It is done when the
recalculated workbook shows the right values and nothing else was damaged. Libraries do not
behave like a spreadsheet application: most of them never rewrite formulas for you.

## Before editing

1. Inspect every sheet, not only the one you will edit: formulas (load with `data_only=False`),
   defined names, charts, data validation, conditional formatting, merged cells and frozen panes.
2. Find what depends on the cells you will change: formulas on other sheets, named ranges,
   chart series, validation and conditional-formatting ranges, totals and subtotals.
3. Check whether "summary" cells are formulas or pasted values. Never assume they recalculate.

## Edits that need extra work

- **Inserting or deleting rows/columns, moving or sorting cells** (`insert_rows`, `delete_rows`,
  `insert_cols`, `move_range`, writing sorted rows back): openpyxl moves cells but does not update
  any reference. Afterwards, rewrite every formula, range, defined name, chart series and
  validation/formatting range that pointed at moved cells, on every sheet.
- **Totals near the end of a table**: a new row inserted directly above a total row is outside
  `SUM(F2:F11)`. Extend the range explicitly.
- **Renaming a sheet**: update every formula and defined name that mentions the old name.
- **Charts**: a chart over a fixed range does not grow with new rows. Update the series references.
- **Validation and conditional formatting**: they apply to fixed ranges; extend them to new rows.
- **Fill-down**: references to a constant cell (a rate, a date) must be absolute (`$B$1`).
- **Lookups**: use exact match (`VLOOKUP(x, range, n, FALSE)`, `MATCH(x, range, 0)`, or `XLOOKUP`).
- **Subtotals**: a grand total over a column that contains subtotal rows double counts. Sum the
  detail rows only, or sum the subtotals.
- **Data types**: write numbers as numbers (strip thousands separators and currency symbols),
  dates as dates, and percentages as fractions (12% is 0.12).
- **Requested formulas**: when results must stay correct after inputs change, write formulas,
  not values computed in Python.
- **Do not round-trip through pandas** (`read_excel` then `to_excel`) to edit a workbook. It
  rewrites the file and drops other sheets, formulas, styles and validation. Edit cells in place.

## Verify before you report

0. Before the first edit, keep a copy of the original: `cp data.xlsx /tmp/original-data.xlsx`.
   After saving, run the audit script that ships with this skill (in its `scripts/` folder):

   ```bash
   python3 <this-skill-dir>/scripts/audit_xlsx.py --before /tmp/original-data.xlsx data.xlsx
   ```

   It needs only openpyxl (and LibreOffice, if installed, to recalculate). Fix every `[error]`.
   Treat every `[warning]` as a question you must answer: either fix it or explain why it is intended.
   The script does not know your task, so a clean result is necessary, not sufficient.
1. Recalculate the saved file with a real engine, for example:

   ```bash
   soffice --headless --convert-to xlsx --outdir /tmp/recalc your-file.xlsx
   ```

   then read `/tmp/recalc/your-file.xlsx` with `openpyxl.load_workbook(..., data_only=True)`.
   (openpyxl writes formulas without results, so reading your own output with `data_only=True`
   shows `None`, not the value.)
2. Compare the key results (totals, lookups, the cells you were asked about) with an independent
   calculation in Python from the raw data.
3. Search all sheets for `#REF!`, `#NAME?`, `#VALUE!`, `#N/A` and `#DIV/0!`.
4. When formulas must stay live, change one input in a copy, recalculate, and confirm the result moves.
5. Compare against the original: sheets, other cells, formats, column widths, validation,
   conditional formatting and charts should be unchanged unless the task asked for a change.

## Reporting

State what you verified and how. If something could not be verified (for example, no spreadsheet
engine was available), say so explicitly instead of claiming the workbook is correct.
