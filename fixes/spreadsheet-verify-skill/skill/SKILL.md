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
1. After saving, run the audit script that ships with this skill (in its `scripts/` folder). It
   recalculates every formula with LibreOffice, if installed, and saves that recalculated copy:

   ```bash
   python3 <this-skill-dir>/scripts/audit_xlsx.py --before /tmp/original-data.xlsx --recalculated /tmp/recalc/data.xlsx data.xlsx
   ```

   It needs only openpyxl (and LibreOffice to recalculate). Fix every `[error]`. Treat every
   `[warning]` as a question you must answer: either fix it or explain why it is intended. A
   finding marked "also in the original file" was there before your edit: change it if the task
   needs it (for example, a total that must now update), otherwise leave it and mention it in your
   report. The script does not know your task, so a clean result is
   necessary, not sufficient. If it reports `recalculation-unavailable`, the values were not
   checked: say so in your report. Exit codes: 0 no `[error]` (warnings may remain), 1 at least
   one `[error]`, 2 invalid command line, 3 a workbook could not be read (missing, truncated,
   password-protected or not an .xlsx/.xlsm file): nothing was checked, so fix or replace the file,
   or report that it could not be verified.
2. Read the results from the recalculated copy with
   `openpyxl.load_workbook("/tmp/recalc/data.xlsx", data_only=True)`. (openpyxl writes formulas
   without results, so reading your own output with `data_only=True` shows `None`, not the value.)
   A plain `soffice --headless --convert-to xlsx` is not a recalculation: with its default
   settings LibreOffice keeps the results already cached in an .xlsx file, so a file with stale
   cached values (for example after editing the XML) still shows the old numbers. To recalculate
   without the script, use a LibreOffice profile that sets `OOXMLRecalcMode` to 0 (always):

   ```bash
   mkdir -p /tmp/lo-recalc/user
   printf '%s\n' '<oor:items xmlns:oor="http://openoffice.org/2001/registry"><item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item></oor:items>' > /tmp/lo-recalc/user/registrymodifications.xcu
   soffice --headless --norestore -env:UserInstallation=file:///tmp/lo-recalc --convert-to xlsx --outdir /tmp/recalc data.xlsx
   ```
3. Compare the key results (totals, lookups, the cells you were asked about) with an independent
   calculation in Python from the raw data.
4. Search all sheets of the recalculated copy for `#REF!`, `#NAME?`, `#VALUE!`, `#N/A` and `#DIV/0!`.
5. When formulas must stay live, change one input in a copy, recalculate it (step 1 or 2), and
   confirm the result moves.
6. Compare against the original: sheets, other cells, formats, column widths, validation,
   conditional formatting and charts should be unchanged unless the task asked for a change.

## What the audit script cannot tell

- It tells moved rows apart by their label in column A or, when that label repeats, by the row's
  constant values (two or more, unique in both files). Rows with a single constant, and duplicate
  rows, are not tracked. Nor are single-cell references on the formula's own sheet without `$`
  (for example a summary below the table that reads `=C2`) when the label in column A repeats or
  is blank. A formula that reads a cell in its own row is taken to belong to that row and move with
  it; it is checked only when it is a summary beside the table: in other columns than a table that
  starts in column A, with an empty column between them, and with a text label of its own, unique in
  its column, that stayed in place (`'Apple price'` next to `G2 =$C$2`); its rows are then told apart
  by the table's own columns. A reference into a table that does not start in column A (a second
  table beside the first) is judged by that table's content, not by the label in column A. After
  sorting or inserting rows, check other such references yourself.
- A formula filled down a column that reads the row next to it on its own sheet without `$` (a
  running balance `=D6+C7`, a change `=B7-B6`) is an `[error]` `reference-meaning-changed` only
  when that row's label in column A and the cells the formula reads there were cleared. It gets
  no finding when its own row or the row it reads is a new row, or when the row it reads holds the
  content that was next to it before (with only deleted rows between them): what refilling the
  column after inserting or deleting rows gives. Otherwise, for example after a sort or when
  repeated rows hide where each row went, it is at most a `[warning]`: a balance refilled after a
  sort reads its new neighbour as intended, so the script cannot call it wrong. (Balances moved by
  an insert and not refilled are reported as an `[error]` `formula-moved-as-text`.) A row
  inserted without its own balance formula gets no finding, although the next balance then reads
  an empty cell, so check that every row of the column has its formula. A one-off formula next to
  the row it reads (a VAT line `=B14*0.2` under the Total) gets none of this leniency.
- A reference that reads another row's content after the rows were reordered is an `[error]`,
  unless it may be meant to read a position; then it is a `[warning]`. That needs a reordering (the
  rows found in both files are in another order; rows only added or removed are none), and a label
  that names a position which the edited table shows the row it reads now holds:
  - its label is the text in the three cells to its left (or the nearest one over formulas), the
    label of the column of formulas it continues ('Top 3' beside the first of three), or one of
    the three nearest texts above it, or above the first cell of a strip of formulas (a header, a
    title). A rank number beside it in a 'Rank', '#', 'No.' or '序号' column counts too, unless its
    line has typed quantities ('Qty', 'Hours', '数量': on an order form or a parts list the number
    is a line number). With no label of its own to its left, a formula beside it in its row that
    reads the same row and qualifies counts;
  - a rank ('#2', '2nd', 'Second place', 'Runner-up', 'Gold', '冠军', '第二名', '倒数第一'), 'Top 3',
    'Last 3' or '前三' need the row at that place in the data, with the table in order of a column
    of numbers, times or dates (the one the label names, else any but an ID or year column: an
    order by SKU or serial number is no ranking); 'highest', 'most', 'max' or '最高' need the highest
    number ('lowest' and the like the lowest; 'Second-highest' the second; 'Biggest drop' the
    lowest) of the column the label names ('most ordered': the Orders column), else of the column
    read; 'top', 'best', 'winner' or '最畅销' need the first row of such an order or the highest
    number, 'bottom' or 'worst' the last or the lowest; 'latest', 'most recent', 'last' or '最新'
    need the newest date of a column of dates ('YYYY-MM', '2026年10月12日', '11月3日', 'FY2025' and
    years under 'Year' count), 'Latest 3' one of the newest three ('earliest' the oldest). A time
    window in a title over a list ('Last 3 months') names the newest three; in the formula's own
    label ('Revenue, last 12 months') it is a field;
  - if the original showed the label held for the row read then (by a column the label names, by
    a column the table was ordered by, or by its dates), it must hold by the same column and in the
    same direction now: 'Top seller' that read the most units must still read the most units, not
    the first name in a list now sorted by name. The original's natural order (IDs, or dates
    running up) and a coincidental highest number show nothing;
  - a formula that pairs rows of two sheets that held the same label in column A (`=Feb!B2-Jan!B2`
    for one store) keeps the `[error]` when they now hold different ones.

  Title, header and note rows and total rows (a 'Total' label, a column total, or typed numbers that
  are the column's sums, at the bottom or the top) are not data rows; a per-row `=SUM(B2:E2)` is
  data, and a customer called 'Total Wine' is a customer. Numbers worked out by simple formulas on
  the row (`=B2*C2`, `=SUM(B2:E2)`) count; other formulas have no numbers to check without
  recalculation, so a claim about them keeps the `[error]`.

  Column names and fixed phrases name no position: the header of the column read, its first words
  or its abbreviation ('Last', 'Last service', 'Last order' for 'Last order date', 'Max. load' for
  'Max load', 'Sales rank', 最低工资 for 最低工资标准), 'First name', '1st quarter', 'last year',
  'first-time', 'first-class', 'third-party', 'last-minute', 'top-line', 'bottom-line', 'per
  second', 'first draft', 'shift leader', '(min)' for minutes, a name with a number ('Warehouse
  #1', 'PO #3'), 第一季度, 第三方, 第一负责人, 倒数天数, 最高学历 and 第二稿. An item's name in column A
  keeps the `[error]` when the labels, the cells left of the formula, the headings above it or its
  sheet's name name it (also in the plural, without spaces, or in part inside Chinese text) and the
  row read now holds another. Other values of that row count only in the formula's own labels (the
  cells left of it, the label of its list, its column's header): another column's text ('Acme'), a
  word only that row has ('Kiwi' for 'Kiwi fruit', 'Gold' for 'Gold plan'), a date as written ('3rd
  March'), a year, a number as '#3' in an ID column, '12号店' or after its column's header ('Store
  12' under 'Store #'), or an ID typed beside the formula. An item named with a position word
  ('Max', 'Gold plan') is therefore taken as named. The script cannot tell a word about history
  from a position ('the first product we bought' gets a warning), does not recognise an item or a
  column named in another language or by another name (a ticker for a company, 'Minimum order qty'
  for 'MOQ', 'Q2 2025' for '2025 Q2'), does not read labels to the right of or below a formula,
  labels that are formulas (`="Top "&D1`) or a card's title more than three texts up, cannot order
  text such as 'P1' or dates as 'DD/MM/YYYY' with no day over 12, and does not carry a position from
  one line of a card to the next or along a row of label and value pairs; the `[error]` stays. If
  the new order was the task, say so in your report.
- Parts that together cover exactly the rows above them, from the first part to the last data row,
  with no gap or overlap and at most three blank rows between (H1 and H2, four quarters), get no
  finding, and neither do totals of whole parts (H1 `=SUM(B2:B7)` beside the quarters, a Year over the
  quarter rows). Rows above the first part are an opening row or an earlier period unless that part
  is the short one (Q1 `=SUM(B3:B4)` beside three-month quarters). Parts that leave rows out, each gap
  shorter than the shortest part (a quarter that skips a month), are one `[error]` at the part next to
  the gap, naming the rows; label and blank rows between sections are not gaps. Parts that overlap are
  an `[error]` naming the rows counted twice. When a total of the whole block stands in the same run
  of totals, above or below the parts, rows left out after the last part are a `[warning]` on the last
  part: H1 2024 and H2 2024 kept below a total extended to 2025 look like a Q4 that misses December.
- A sum or other aggregate that is the whole formula and starts below the first data row of its
  column's block (`=SUM(B3:B13)` over data from row 2) is a `[warning]`, since a first row such as an
  opening balance may be left out on purpose; not when an aggregate of the same function in the same
  block of totals covers those rows and its own (a 'Q4 2024' below the year's total), or when parts
  before it sum the rows above it (H2 below H1). It is an `[error]` when a part in that block includes
  both its first row and the rows above it while this aggregate contains another part (a 'Total' below
  2024 / 2025 parts that leaves out January), or, with `--before`, when the original's formula in the
  same cell, or with the same label in column A, covered the left-out rows and they hold the same
  values. Header rows are not data rows: row 1, the first filled cell of each column, and a row with
  no label in column A above rows that have one (years below the column titles); nor are a row
  computed only from totals above it (a Variance) and a rounded total (`=ROUND(SUM(B2:B13),0)`).
- A shorter aggregate below a total that covers more rows (an Average, a 'Q4 2024' sum) is a
  `[warning]`. It is an `[error]` only with `--before`, when in the original it covered the same
  rows (the same first and last row) as another aggregate above it with a different formula (two
  identical formulas cannot be told apart) and the total that now covers the new rows starts at the
  same row. A part that starts lower, such as a quarter or half-year kept below an all-time total,
  stays a `[warning]`, because correct edits keep such parts too. With `--before`, a total left at
  its old range and followed by a new sum of only the added rows is a `[warning]` for the same
  reason (it has the shape of a correct split such as 2024 / 2025), unless a total of all the rows
  sits in the same block of total rows.
- A new sheet that holds at least half of a removed sheet's constants counts as that sheet
  renamed; formulas are compared only for a sheet that held no constants, so a sheet renamed with
  its values cleared and its formulas kept is reported as removed. A chart sheet counts as renamed
  when a new chart sheet has the same title or a series whose values (the y values of a scatter
  chart) come from at least one of the same columns or defined names (a name that stands for a
  range counts as that range), even when the new chart shows only one of the old chart's columns.
  An old chart with neither a title nor a series over a cell range (for example a series over an
  `OFFSET` name) is a `[warning]` when a new chart sheet is left over, since nothing tells the two
  apart. A chart sheet in a file re-saved by LibreOffice 24.2 becomes a worksheet and is reported
  as removed.
- Known misreadings: a total written as `=SUM(B2:B4)+SUM(B5:B7)`, and half-year sums four or more
  blank rows below the data, are `[error]`s; conditional formatting on column A that stops above a
  'Total' label is a `[warning]`. Known misses: months appended below half-year sums that have no
  total are not reported; a part of an exact split that starts late (2024 `=SUM(B3:B13)` beside a
  2025 part) is reported only with `--before`; and 5 of 150 random wrong `insert_rows` edits of
  invoice tables with repeated values passed clean in our tests.

## Reporting

State what you verified and how. If something could not be verified (for example, no spreadsheet
engine was available), say so explicitly instead of claiming the workbook is correct.
