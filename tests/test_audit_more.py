"""The audit shipped with spreadsheet-verify-skill: no errors on ordinary correct edits, errors kept
on real breakage, recalculation that is forced and fails softly, unreadable or oversized inputs
handled without a traceback, and the skill installed whole."""

import os
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import openpyxl
import pytest

from failbook import audit as audit_mod
from failbook.audit import audit, check_errors, main
from failbook.case import load_case, repo_root
from failbook.fixes import load_fix
from failbook.runner import PRESETS

from conftest import needs_libreoffice

ROOT = repo_root()
SKILL = ROOT / "fixes" / "spreadsheet-verify-skill"


def _levels(findings, check):
    return [f.level for f in findings if f.check == check]


def _errors(findings):
    return [f for f in findings if f.level == "error"]


def _stats_sheet(path, extra=None):
    """B2:B10 data with a Total / Average / Max block below it (a common layout)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws["A1"], ws["B1"] = "Month", "Revenue"
    for r in range(2, 11):
        ws.cell(r, 1, f"M{r - 1}")
        ws.cell(r, 2, 100 * r)
    ws["A11"], ws["B11"] = "Total", "=SUM(B2:B10)"
    ws["A12"], ws["B12"] = "Average", "=AVERAGE(B2:B10)"
    ws["A13"], ws["B13"] = "Max", "=MAX(B2:B10)"
    ws["A14"], ws["B14"] = "Share of max", "=B11/B13"
    for cell, value in (extra or {}).items():
        ws[cell] = value
    wb.save(path)
    return path


# -- range-omits-adjacent-cells ----------------------------------------------------------------
def test_stacked_summary_rows_are_not_omitted_data(tmp_path):
    before = _stats_sheet(tmp_path / "before.xlsx")
    after = _stats_sheet(tmp_path / "after.xlsx")
    assert not _errors(audit(after, before, recalc=False))
    assert not _errors(audit(after, recalc=False))


def test_data_row_between_range_and_total_is_still_an_error(tmp_path):
    before = _stats_sheet(tmp_path / "before.xlsx")
    # A new month typed in row 11, the summary block moved down by one row, ranges not extended.
    after = _stats_sheet(tmp_path / "after.xlsx", {
        "A11": "M10", "B11": 1200, "A12": "Total", "B12": "=SUM(B2:B10)", "A13": "Average",
        "B13": "=AVERAGE(B2:B10)", "A14": "Max", "B14": "=MAX(B2:B10)", "A15": "Share of max", "B15": "=B12/B14"})
    found = audit(after, before, recalc=False)
    assert _levels(found, "range-omits-adjacent-cells")[:1] == ["error"]
    assert any("rows 11-11" in f.message for f in found if f.check == "range-omits-adjacent-cells")


def test_symptom_already_in_the_original_is_a_warning(tmp_path):
    # The original already has a value in row 11 that the total skips; an unrelated edit keeps it.
    before = _stats_sheet(tmp_path / "before.xlsx", {"A11": "Note", "B11": 5, "A12": "Total", "B12": "=SUM(B2:B10)",
                                                     "A13": None, "B13": None, "A14": None, "B14": None})
    after = _stats_sheet(tmp_path / "after.xlsx", {"A11": "Note", "B11": 5, "A12": "Total", "B12": "=SUM(B2:B10)",
                                                   "A13": None, "B13": None, "A14": None, "B14": None, "A1": "Months"})
    assert _levels(audit(before, recalc=False), "range-omits-adjacent-cells") == ["error"]
    found = [f for f in audit(after, before, recalc=False) if f.check == "range-omits-adjacent-cells"]
    assert [f.level for f in found] == ["warning"] and "also in the original file" in found[0].message


# -- sheet-removed / stale-sheet-reference -----------------------------------------------------
def _xlsx006(tmp_path, *, append_row, update_refs):
    case = load_case("xlsx-006")
    name = case.outputs[0]
    before = case.input_dir / name
    wb = openpyxl.load_workbook(before)
    old, new = "Data", "Sales 2024"  # the rename the task asks for
    assert old in wb.sheetnames
    wb[old].title = new
    if append_row:
        ws = wb[new]
        last = ws.max_row
        for col in range(1, ws.max_column + 1):
            v = ws.cell(last, col).value
            if not (isinstance(v, str) and v.startswith("=")):
                ws.cell(last + 1, col, v)
    if update_refs:
        pattern = re.compile(r"(?<![A-Za-z0-9_.'])" + re.escape(old) + r"!")
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and c.value.startswith("="):
                        c.value = pattern.sub(f"'{new}'!", c.value.replace(f"'{old}'!", f"'{new}'!"))
        for dn, d in list(wb.defined_names.items()):
            d.attr_text = pattern.sub(f"'{new}'!", str(d.attr_text).replace(f"'{old}'!", f"'{new}'!"))
    after = tmp_path / name
    wb.save(after)
    return before, after


def test_renamed_and_edited_sheet_is_not_reported_as_removed(tmp_path):
    before, after = _xlsx006(tmp_path, append_row=True, update_refs=True)
    found = audit(after, before, recalc=False)
    assert "sheet-removed" not in {f.check for f in found}
    assert not _errors(found)


def test_renamed_sheet_with_stale_references_is_still_an_error(tmp_path):
    before, after = _xlsx006(tmp_path, append_row=True, update_refs=False)
    found = audit(after, before, recalc=False)
    assert "sheet-removed" not in {f.check for f in found}
    assert "error" in _levels(found, "stale-sheet-reference")


def test_sheet_replaced_by_unrelated_content_is_removed(tmp_path):
    case = load_case("xlsx-006")
    before = case.input_dir / case.outputs[0]
    wb = openpyxl.load_workbook(before)
    first = wb.sheetnames[0]
    del wb[first]
    ws = wb.create_sheet("Sales 2024", 0)
    ws["A1"] = "something else entirely"
    after = tmp_path / "report.xlsx"
    wb.save(after)
    assert _levels(audit(after, before, recalc=False), "sheet-removed") == ["error"]


# -- reference-meaning-changed -----------------------------------------------------------------
def test_renamed_header_is_a_warning_and_shifted_column_an_error(tmp_path):
    case = load_case("xlsx-005")
    name = case.outputs[0]
    before = case.input_dir / name
    wb = openpyxl.load_workbook(before)
    ws = wb["Lines"]
    ws["B1"] = f"{ws['B1'].value} (EUR)"
    after = tmp_path / name
    wb.save(after)
    found = audit(after, before, recalc=False)
    assert not _errors(found)
    assert _levels(found, "reference-label-changed") == ["warning"]
    # Converted in place (new header, new values, nothing moved): a question, not a verdict.
    for r in range(2, ws.max_row + 1):
        if isinstance(ws.cell(r, 2).value, (int, float)):
            ws.cell(r, 2, round(ws.cell(r, 2).value * 0.9, 2))
    wb.save(after)
    found = audit(after, before, recalc=False)
    assert not _errors(found)
    assert _levels(found, "reference-meaning-changed") == ["warning"]
    # The committed failure artifact inserts a column, so the same formulas now read other data.
    assert "error" in _levels(audit(case.failure_dir / name, before, recalc=False), "reference-meaning-changed")


def test_renamed_row_label_is_a_warning_and_deleted_rows_an_error(tmp_path):
    case = load_case("xlsx-014")
    name = case.outputs[0]
    before = case.input_dir / name
    wb = openpyxl.load_workbook(before)
    ws = wb[wb.sheetnames[0]]
    total_row = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "Total")
    ws.cell(total_row, 1, "Grand total")
    after = tmp_path / name
    wb.save(after)
    found = audit(after, before, recalc=False)
    assert not _errors(found)
    assert "warning" in _levels(found, "reference-label-changed")
    assert "error" in _levels(audit(case.failure_dir / name, before, recalc=False), "reference-meaning-changed")


# -- correct edits from the second review: no error; the wrong variants still get one ----------
def _book(path, rows, title="Sheet", extra=None):
    """A one-sheet workbook from a list of rows; ``extra`` maps cells to values."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title
    for row in rows:
        ws.append(row)
    for cell, value in (extra or {}).items():
        ws[cell] = value
    wb.save(path)
    return path


def _edit(src, dest, change):
    wb = openpyxl.load_workbook(src)
    change(wb)
    wb.save(dest)
    return dest


def _findings(after, before=None):
    return [f for f in audit(after, before, recalc=False) if f.check != "recalculation-skipped"]


LEDGER = [("Opening", 1000), ("Rent", -400), ("Invoice 17", 650), ("Fuel", -60), ("Invoice 18", 300),
          ("Phone", -45), ("Bank fee", -5), ("Invoice 19", 820)]


def _ledger(path, items=None):
    """A cash ledger with a running balance: D_r = D_(r-1) + C_r. ``items`` are (date, text, amount)."""
    items = items or [(f"2024-03-{r:02d}", d, a) for r, (d, a) in enumerate(LEDGER, start=2)]
    rows = [["Date", "Description", "Amount", "Balance"]]
    rows += [[day, d, a, "=C2" if r == 2 else f"=D{r - 1}+C{r}"] for r, (day, d, a) in enumerate(items, start=2)]
    return _book(path, rows, "Ledger")


def test_workbook_with_a_chart_sheet_is_compared(tmp_path):
    from openpyxl.chart import BarChart, Reference

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Month", "Revenue"])
    for i in range(1, 7):
        ws.append([f"M{i}", 100 * i])
    ws["A8"], ws["B8"] = "Total", "=SUM(B2:B7)"
    chart = BarChart()
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=7), titles_from_data=True)
    wb.create_chartsheet("Chart").add_chart(chart)
    before = tmp_path / "before.xlsx"
    wb.save(before)
    after = _edit(before, tmp_path / "after.xlsx", lambda w: w["Data"].__setitem__("B3", 250))
    assert _findings(after, before) == []
    assert main([str(after), "--before", str(before), "--no-recalc"]) == 0
    gone = _edit(before, tmp_path / "gone.xlsx", lambda w: w.remove(w["Chart"]))
    assert [(f.level, f.check, f.where) for f in _findings(gone, before)] == [("error", "sheet-removed", "Chart")]


def test_sort_that_breaks_single_cell_references_is_an_error(tmp_path):
    rows = [["Product", "Category", "Price", "Unit", "Supplier"]]
    rows += [[p, "Fruit", x, "kg", "Acme"] for p, x in [("Apple", 2.1), ("Banana", 1.2), ("Kiwi", 3.4), ("Mango", 2.9)]]
    before = _book(tmp_path / "before.xlsx", rows, "Prices")
    wb = openpyxl.load_workbook(before)
    summary = wb.create_sheet("Summary")
    summary["A1"], summary["B1"] = "Apple price", "=Prices!C2"
    wb.save(before)

    def sort(w):  # rows written back sorted by price; =Prices!C2 left alone, so it now shows Kiwi's price
        ws = w["Prices"]
        data = sorted(([c.value for c in r] for r in ws.iter_rows(min_row=2)), key=lambda r: -r[2])
        for r, values in enumerate(data, start=2):
            for col, v in enumerate(values, start=1):
                ws.cell(r, col, v)

    found = _findings(_edit(before, tmp_path / "after.xlsx", sort), before)
    assert [(f.level, f.check, f.where) for f in found] == [("error", "reference-meaning-changed", "Summary!B1")]
    assert "'Apple' is now in row 4" in found[0].message
    assert not any("stayed the same" in f.message for f in found)


@pytest.mark.parametrize("edit", ["append", "sort", "title"])
def test_renamed_sheet_that_was_also_extended_sorted_or_titled_is_a_rename(tmp_path, edit):
    reps = ["Ana", "Ben", "Cai", "Dee", "Eli", "Fay", "Gus", "Hal", "Ivy", "Jon", "Kim", "Lou"]
    rows = [["Region", "Rep", "Revenue"]] + [["North" if i % 2 else "South", r, 1000 + 137 * ((i * 7) % 12)]
                                             for i, r in enumerate(reps)]
    before = _book(tmp_path / "before.xlsx", rows, "Data")
    wb = openpyxl.load_workbook(before)
    wb.create_sheet("Summary")["B2"] = "=SUM(Data!C2:C13)"
    wb.save(before)

    def change(w):
        ws = w["Data"]
        ws.title = "Sales 2024"
        first = 2
        if edit == "append":
            for i in range(15):
                ws.append(["North" if i % 3 else "South", f"New{i}", 900 + 11 * i])
        elif edit == "sort":
            data = sorted(([c.value for c in r] for r in ws.iter_rows(min_row=2)), key=lambda r: -r[2])
            for r, values in enumerate(data, start=2):
                for col, v in enumerate(values, start=1):
                    ws.cell(r, col, v)
        else:
            ws.insert_rows(1)
            ws["A1"] = "Sales by rep, 2024"
            first = 3
        w["Summary"]["B2"] = f"=SUM('Sales 2024'!C{first}:C{ws.max_row})"

    found = _findings(_edit(before, tmp_path / "after.xlsx", change), before)
    assert "sheet-removed" not in {f.check for f in found}
    assert not _errors(found)


def test_inserted_column_or_spacer_row_keeps_correct_unchanged_formulas(tmp_path):
    rows = [["Item", "Qty", "Price", "Total"]]
    rows += [[it, q, p, f"=B{r}*C{r}"] for r, (it, q, p) in enumerate([("Pen", 10, 1.5), ("Pad", 4, 3.2), ("Ink", 2, 12.0)], start=2)]
    before = _book(tmp_path / "orders.xlsx", rows + [["Total", None, None, "=SUM(D2:D4)"]], "Orders")

    def insert_column(w):  # a Notes column before Total: Excel keeps =B2*C2, now in column E
        ws = w.active
        ws.insert_cols(4)
        ws["D1"], ws["E5"] = "Notes", "=SUM(E2:E4)"

    assert _findings(_edit(before, tmp_path / "inscol.xlsx", insert_column), before) == []
    sales = _book(tmp_path / "sales.xlsx", [["Month", "Revenue"]] + [[f"M{r - 1}", 100 * r] for r in range(2, 11)],
                  "Sales", {"A11": "Total", "B11": "=SUM(B2:B10)"})
    spaced = _edit(sales, tmp_path / "gap.xlsx", lambda w: w.active.insert_rows(11))  # Excel keeps =SUM(B2:B10)
    assert _findings(spaced, sales) == []


def test_formulas_moved_with_the_rows_they_read_are_still_an_error(tmp_path):
    before = _ledger(tmp_path / "before.xlsx")

    def insert(w):  # openpyxl moves the formulas down as text; =D5+C6 now sits in row 7 and reads row 6
        ws = w.active
        ws.insert_rows(6)
        ws["A6"], ws["B6"], ws["C6"] = "2024-03-05", "Refund", 25

    found = _findings(_edit(before, tmp_path / "after.xlsx", insert), before)
    assert [(f.level, f.check) for f in found] == [("error", "formula-moved-as-text")]
    case = load_case("xlsx-004")
    name = case.outputs[0]
    assert "error" in _levels(audit(case.failure_dir / name, case.input_dir / name, recalc=False), "formula-moved-as-text")


def test_running_balance_refilled_after_an_inserted_row_is_clean(tmp_path):
    before = _ledger(tmp_path / "before.xlsx")

    def insert(w):  # the correct final state: every row r has =D(r-1)+C(r)
        ws = w.active
        ws.insert_rows(6)
        ws["A6"], ws["B6"], ws["C6"] = "2024-03-05", "Refund", 25
        for r in range(3, ws.max_row + 1):
            ws.cell(r, 4, f"=D{r - 1}+C{r}")

    assert _findings(_edit(before, tmp_path / "after.xlsx", insert), before) == []
    # The same final state after deleting a transaction, or after inserting two at different places
    # (two rows share a date there, so the dates alone do not say where each row came from).
    items = [(f"2024-03-{r:02d}", d, a) for r, (d, a) in enumerate(LEDGER, start=2)]
    deleted = _ledger(tmp_path / "deleted.xlsx", items[:4] + items[5:])
    assert _findings(deleted, before) == []
    two = items[:3] + [("2024-03-04", "Refund", 25)] + items[3:6] + [("2024-03-07", "Deposit", 90)] + items[6:]
    assert _findings(_ledger(tmp_path / "two.xlsx", two), before) == []
    # An absolute row ($5) in the same layout does not follow the rows: still an error.
    rows = [["Item", "Base", "Qty", "Cost"]] + [[f"I{r}", 10 * r, r, f"=B$5*C{r}"] for r in range(2, 10)]
    anchored = _book(tmp_path / "anchored.xlsx", rows, "Calc")

    def insert_above(w):
        ws = w.active
        ws.insert_rows(3)
        ws["A3"], ws["B3"], ws["C3"] = "new", 7, 1
        for r in range(2, 11):
            ws.cell(r, 4, f"=B$5*C{r}")

    found = _findings(_edit(anchored, tmp_path / "anchored-after.xlsx", insert_above), anchored)
    assert _errors(found) and {f.check for f in _errors(found)} == {"reference-meaning-changed"}


def test_half_year_sums_below_a_total_are_not_omitted_data(tmp_path):
    rows = [["Month", "Revenue"]] + [[f"2024-{r - 1:02d}", 100 + 10 * r] for r in range(2, 14)]
    before = _book(tmp_path / "before.xlsx", rows, "Sales", {"A14": "Total", "B14": "=SUM(B2:B13)"})
    halves = _edit(before, tmp_path / "halves.xlsx", lambda w: [w.active.__setitem__(k, v) for k, v in
                                                                 {"A15": "H1", "B15": "=SUM(B2:B7)",
                                                                  "A16": "H2", "B16": "=SUM(B8:B13)"}.items()])
    assert _findings(halves, before) == []
    # A part that does not split the total exactly is a question, not a verdict.
    first_half = _edit(before, tmp_path / "h1.xlsx", lambda w: w.active.__setitem__("B15", "=SUM(B2:B7)"))
    assert _levels(_findings(first_half, before), "range-omits-adjacent-cells") == ["warning"]
    # The same split above the Year total, or with no total at all, is clean too.
    above = _book(tmp_path / "above.xlsx", rows, "Sales", {"A14": "H1", "B14": "=SUM(B2:B7)", "A15": "H2",
                                                            "B15": "=SUM(B8:B13)", "A16": "Year", "B16": "=SUM(B2:B13)"})
    assert _findings(above, before) == []
    plain = _book(tmp_path / "plain.xlsx", rows, "Sales")
    alone = _book(tmp_path / "alone.xlsx", rows, "Sales", {"A14": "H1", "B14": "=SUM(B2:B7)", "A15": "H2", "B15": "=SUM(B8:B13)"})
    assert _findings(alone, plain) == [] and _findings(alone) == []
    # Two months added below the data and the halves left as they were: still an error.
    grown = rows + [["2025-01", 300], ["2025-02", 310]]
    stale = _book(tmp_path / "stale.xlsx", grown, "Sales", {"A16": "H1", "B16": "=SUM(B2:B7)", "A17": "H2", "B17": "=SUM(B8:B13)"})
    assert _levels(_findings(stale, plain), "range-omits-adjacent-cells") == ["error", "error"]


def test_partial_aggregate_left_behind_by_an_extended_total_is_an_error(tmp_path):
    rows = [["Month", "Revenue"]] + [[f"2024-{r - 1:02d}", 100 + 10 * r] for r in range(2, 14)]
    before = _book(tmp_path / "before.xlsx", rows, "Sales", {"A14": "Total", "B14": "=SUM(B2:B13)",
                                                              "A15": "Average", "B15": "=AVERAGE(B2:B13)"})

    def add_two_months(average):  # rows inserted above the total, as Excel would; the Total extended
        def change(w):
            ws = w.active
            ws.insert_rows(14, 2)
            ws["A14"], ws["B14"], ws["A15"], ws["B15"] = "2025-01", 300, "2025-02", 310
            ws["B16"], ws["B17"] = "=SUM(B2:B15)", average
        return change

    stale = _edit(before, tmp_path / "stale.xlsx", add_two_months("=AVERAGE(B2:B13)"))
    found = [f for f in _findings(stale, before) if f.check == "range-omits-adjacent-cells"]
    assert [(f.level, f.where) for f in found] == [("error", "Sales!B17")]
    assert "in the original file both covered rows 2-13 (the total was in row 14: =SUM(B2:B13))" in found[0].message
    # Without the original nothing shows that the shorter range is stale: a warning.
    assert _levels(_findings(stale), "range-omits-adjacent-cells") == ["warning"]
    assert _findings(_edit(before, tmp_path / "ok.xlsx", add_two_months("=AVERAGE(B2:B15)")), before) == []


def test_chart_that_leaves_out_the_total_row_is_clean(tmp_path):
    from openpyxl.chart import BarChart, Reference

    def chart_book(path, months, last):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Revenue"
        ws.append(["Month", "Revenue"])
        for i, m in enumerate(months):
            ws.append([m, 1000 + 75 * i])
        ws.cell(len(months) + 2, 1, "Total")
        ws.cell(len(months) + 2, 2, f"=SUM(B2:B{len(months) + 1})")
        chart = BarChart()
        chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=last), titles_from_data=True)
        ws.add_chart(chart, "D2")
        wb.save(path)
        return path

    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
    assert _findings(chart_book(tmp_path / "ok.xlsx", months, 7)) == []
    # Jul and Aug typed above the total, chart not extended: the warning stops at the last month.
    found = _findings(chart_book(tmp_path / "short.xlsx", months + ["Jul", "Aug"], 7))
    assert [f.level for f in found] == ["warning"] and "data continues to row 9" in found[0].message


# -- third review: wrong edits that had stopped failing, correct ones still flagged, slow scans ---
def _invoice(path):
    rows = [["Service", "Hours", "Rate", "VAT", "Amount"]]
    rows += [[s, 1, rate, 0.2, f"=B{r}*C{r}*(1+D{r})"]
             for r, (s, rate) in enumerate([("Audit", 120), ("Tax return", 95), ("Payroll", 80), ("Advice", 150)], start=2)]
    return _book(path, rows, "Invoice")


def test_moved_formulas_are_an_error_when_the_columns_they_read_repeat_values(tmp_path):
    before = _invoice(tmp_path / "before.xlsx")

    def insert(rewrite):
        def change(w):  # a new first line; Hours (all 1) and VAT (all 0.2) read the same at either row
            ws = w.active
            ws.insert_rows(2)
            ws["A2"], ws["B2"], ws["C2"], ws["D2"], ws["E2"] = "Bookkeeping", 1, 60, 0.2, "=B2*C2*(1+D2)"
            for r in range(3, 7) if rewrite else ():
                ws.cell(r, 5, f"=B{r}*C{r}*(1+D{r})")
        return change

    found = _findings(_edit(before, tmp_path / "wrong.xlsx", insert(False)), before)
    assert [(f.level, f.check, f.where) for f in found] == [("error", "formula-moved-as-text", "Invoice!E3")]
    assert _findings(_edit(before, tmp_path / "right.xlsx", insert(True)), before) == []


def test_whole_column_formulas_after_an_inserted_column_are_read_once_per_range(tmp_path, monkeypatch):
    rows = [["Customer", "Amount", "Date", "Customer total"]]
    rows += [[f"C{r % 37}", r % 101, f"2024-01-{r % 28 + 1:02d}", f"=SUMIF(A:A,A{r},B:B)"] for r in range(2, 2002)]
    before = _book(tmp_path / "before.xlsx", rows, "Orders")

    def insert_notes(w):  # what Excel leaves: the formulas in column E keep reading A and B
        ws = w.active
        ws.insert_cols(4)
        ws["D1"] = "Notes"

    after = _edit(before, tmp_path / "after.xlsx", insert_notes)
    reads = {"n": 0}
    value = audit_mod._value

    def counting(ws, row, col):
        reads["n"] += 1
        return value(ws, row, col)

    monkeypatch.setattr(audit_mod, "_value", counting)
    assert _findings(after, before) == []
    # The quadratic scan read up to 2,000 cells for each of the 2,000 moved formulas.
    assert reads["n"] < 200_000, reads["n"]


def test_formula_only_sheet_renamed_and_given_a_note_is_a_rename(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Region", "Revenue"])
    for i in range(8):
        ws.append(["North" if i % 2 else "South", 100 + 7 * i])
    calc = wb.create_sheet("Calc")
    calc["A1"], calc["B1"] = "=Data!A1", "=Data!B1"
    calc["A2"], calc["B2"] = '="North"', '=SUMIF(Data!A2:A9,"North",Data!B2:B9)'
    calc["A3"], calc["B3"] = '="South"', '=SUMIF(Data!A2:A9,"South",Data!B2:B9)'
    calc["A4"], calc["B4"] = '="Total"', "=SUM(B2:B3)"
    wb.create_sheet("Report").append(["Grand total", "=Calc!B4"])
    before = tmp_path / "before.xlsx"
    wb.save(before)

    def rename(note, update=True):
        def change(w):
            w["Calc"].title = "Calculations"
            w["Report"]["B1"] = "=Calculations!B4" if update else "=Calc!B4"
            if note:
                w["Calculations"]["D1"] = "Source: Data sheet"
        return change

    for note in (False, True):
        assert _findings(_edit(before, tmp_path / f"after{note}.xlsx", rename(note)), before) == []
    found = _findings(_edit(before, tmp_path / "stale.xlsx", rename(True, update=False)), before)
    assert [(f.level, f.check) for f in found] == [("error", "stale-sheet-reference")]


def test_total_written_without_an_aggregate_ends_the_data_for_charts_and_highlighting(tmp_path):
    from openpyxl.chart import BarChart, Reference
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import PatternFill

    def book(path, total, months=6):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Revenue"
        ws.append(["Month", "Revenue"])
        for i in range(months):
            ws.append([f"M{i + 1}", 1000 + 75 * i])
        ws.cell(months + 2, 1, "Total")
        ws.cell(months + 2, 2, total)
        chart = BarChart()
        chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=7), titles_from_data=True)
        ws.add_chart(chart, "D2")
        ws.conditional_formatting.add("A2:B7", CellIsRule(operator="greaterThan", formula=["1200"],
                                                          fill=PatternFill("solid", fgColor="FFFF00")))
        wb.save(path)
        return path

    assert _findings(book(tmp_path / "plus.xlsx", "=B2+B3+B4+B5+B6+B7")) == []
    assert _findings(book(tmp_path / "sum.xlsx", "=SUM(B2:B7)")) == []  # the CF covers column A's 'Total' row
    found = _findings(book(tmp_path / "short.xlsx", "=B2+B3+B4+B5+B6+B7+B8+B9", months=8))
    assert [f.level for f in found] == ["warning", "warning"]
    assert all("data continues to row 9" in f.message for f in found)


def test_chart_sheet_counts_as_renamed_only_when_it_shows_the_same_chart(tmp_path):
    from openpyxl.chart import BarChart, Reference

    def book(path, sheets, rows=7):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["Month", "Revenue", "Cost"])
        for i in range(1, rows):
            ws.append([f"M{i}", 100 * i, 70 * i])
        for name, col in sheets:
            chart = BarChart()
            chart.add_data(Reference(ws, min_col=col, min_row=1, max_row=rows), titles_from_data=True)
            wb.create_chartsheet(name).add_chart(chart)
        wb.save(path)
        return path

    before = book(tmp_path / "before.xlsx", [("Revenue chart", 2)])
    assert _findings(book(tmp_path / "renamed.xlsx", [("Revenue", 2)], rows=9), before) == []
    found = _findings(book(tmp_path / "swap.xlsx", [("Cost chart", 3)]), before)
    assert [(f.level, f.check, f.where) for f in found] == [("error", "sheet-removed", "Revenue chart")]
    assert "'Cost chart'" in found[0].message


def test_running_total_column_is_scanned_for_double_counting_in_linear_work(tmp_path, monkeypatch):
    rows = [["Day", "Sales", "Cumulative"]] + [[f"D{r - 1}", (r * 37) % 101, f"=SUM(B$2:B{r})"] for r in range(2, 3002)]
    path = _book(tmp_path / "daily.xlsx", rows, "Daily")
    reads = {"n": 0}
    value = audit_mod._value

    def counting(ws, row, col):
        reads["n"] += 1
        return value(ws, row, col)

    monkeypatch.setattr(audit_mod, "_value", counting)
    assert _findings(path) == []
    assert reads["n"] < 100_000, reads["n"]  # each of the 3,000 sums visited its whole range before
    # A grand total over a column that holds its own subtotals is still found.
    rows = [["Item", "Amount"], ["a", 1], ["b", 2], ["Sub 1", "=SUM(B2:B3)"], ["c", 3], ["Sub 2", "=SUM(B5:B5)"],
            ["Total", "=SUM(B2:B6)"]]
    found = _findings(_book(tmp_path / "subtotals.xlsx", rows, "S"))
    assert [(f.level, f.check, f.where) for f in found] == [("error", "double-counting", "S!B7")]


def test_sort_under_a_repeated_category_label_is_found_by_the_row_values(tmp_path):
    rows = [["Category", "Product", "Price", "Unit", "Supplier"]]
    rows += [["Fruit", p, x, "kg", "Acme"] for p, x in [("Apple", 2.1), ("Banana", 1.2), ("Kiwi", 3.4), ("Mango", 2.9)]]
    before = _book(tmp_path / "before.xlsx", rows, "Prices")
    wb = openpyxl.load_workbook(before)
    wb.create_sheet("Summary").append(["Apple price", "=Prices!C2"])
    wb.save(before)

    def sort(w):  # =Prices!C2 left alone: it now shows Kiwi's price
        ws = w["Prices"]
        data = sorted(([c.value for c in r] for r in ws.iter_rows(min_row=2)), key=lambda r: -r[2])
        for r, values in enumerate(data, start=2):
            for col, v in enumerate(values, start=1):
                ws.cell(r, col, v)

    found = _findings(_edit(before, tmp_path / "after.xlsx", sort), before)
    assert [(f.level, f.check, f.where) for f in found] == [("error", "reference-meaning-changed", "Summary!B1")]
    assert "are now in row 4" in found[0].message
    # Updating the reference, appending a product or changing a price in place stays clean.
    fixed = _edit(tmp_path / "after.xlsx", tmp_path / "fixed.xlsx", lambda w: w["Summary"].__setitem__("B1", "=Prices!C4"))
    assert _findings(fixed, before) == []
    assert _findings(_edit(before, tmp_path / "more.xlsx", lambda w: w["Prices"].append(["Fruit", "Pear", 1.8, "kg", "Acme"])),
                     before) == []
    assert _findings(_edit(before, tmp_path / "price.xlsx", lambda w: w["Prices"].__setitem__("C2", 2.3)), before) == []


# -- fourth review: the round-3 rules tightened where correct edits share the shape ------------
def _chart_sheet_book(path, sheet, series, title=None, kind="bar", data="Data"):
    """A data sheet (M, R, C, P) and one chart sheet ``sheet`` whose series read the columns ``series``."""
    from openpyxl.chart import BarChart, Reference, ScatterChart, Series

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = data
    ws.append(["M", "R", "C", "P"])
    for i in range(1, 8):
        ws.append([i, i * i, 7 - i, 2 * i])
    if kind == "scatter":  # scatter series keep their values in yVal (and xVal), not in val
        chart = ScatterChart()
        for col in series:
            chart.series.append(Series(Reference(ws, min_col=col, min_row=2, max_row=8), Reference(ws, min_col=1, min_row=2, max_row=8)))
    else:
        chart = BarChart()
        for col in series:
            chart.add_data(Reference(ws, min_col=col, min_row=1, max_row=8), titles_from_data=True)
    if title:
        chart.title = title
    wb.create_chartsheet(sheet).add_chart(chart)
    wb.save(path)
    return path


def test_renamed_chart_sheet_matches_scatter_series_and_added_series(tmp_path):
    for kind in ("scatter", "bar"):
        before = _chart_sheet_book(tmp_path / f"{kind}_b.xlsx", "Chart1", [2], kind=kind)
        assert _findings(_chart_sheet_book(tmp_path / f"{kind}_ren.xlsx", "Growth", [2], kind=kind), before) == []
    # An untitled chart renamed and given a Cost series; a titled one renamed, retitled and extended.
    before = _chart_sheet_book(tmp_path / "b.xlsx", "Chart1", [2])
    assert _findings(_chart_sheet_book(tmp_path / "add.xlsx", "Revenue and cost", [2, 3]), before) == []
    before = _chart_sheet_book(tmp_path / "bt.xlsx", "Revenue chart", [2], title="Revenue")
    assert _findings(_chart_sheet_book(tmp_path / "title.xlsx", "Sales chart", [2, 3], title="Revenue and cost"), before) == []
    # Another chart in its place (no title, no column in common) is still a removed chart sheet.
    found = _findings(_chart_sheet_book(tmp_path / "other.xlsx", "Profit chart", [4], title="Profit"), before)
    assert [(f.level, f.check, f.where) for f in found] == [("error", "sheet-removed", "Revenue chart")]
    # The skill's fallback chart reader (no failbook package) reads scatter values too.
    from failbook.xlsx import charts

    scatter = _chart_sheet_book(tmp_path / "scatter_b.xlsx", "Chart1", [2], kind="scatter")
    assert [s["val"] for c in audit_mod._charts_fallback(scatter) for s in c["series"]] == ["'Data'!$B$2:$B$8"]
    assert [s["val"] for c in charts(scatter) for s in c["series"]] == ["'Data'!$B$2:$B$8"]


def test_part_kept_below_an_extended_total_is_a_warning_and_a_stale_full_range_an_error(tmp_path):
    months = [f"2024-{m:02d}" for m in range(1, 13)] + ["2025-01", "2025-02"]

    def sales(path, extra, n=12):
        rows = [["Month", "Revenue"]] + [[months[r - 2], 100 + 10 * r] for r in range(2, n + 2)]
        return _book(path, rows, "Sales", extra)

    # 'Q4 2024' and 'H1/H2 2024' keep their range under an all-time total: correct edits do that too.
    before = sales(tmp_path / "qb.xlsx", {"A14": "Total", "B14": "=SUM(B2:B13)", "A15": "Q4 2024", "B15": "=SUM(B11:B13)"})
    q4 = sales(tmp_path / "q4.xlsx", {"A16": "Total", "B16": "=SUM(B2:B15)", "A17": "Q4 2024", "B17": "=SUM(B11:B13)"}, 14)
    assert [f.level for f in _findings(q4, before)] == ["warning"]
    before = sales(tmp_path / "hb.xlsx", {"A14": "Total", "B14": "=SUM(B2:B13)", "A15": "H1 2024", "B15": "=SUM(B2:B7)",
                                          "A16": "H2 2024", "B16": "=SUM(B8:B13)"})
    halves = sales(tmp_path / "h.xlsx", {"A16": "Total", "B16": "=SUM(B2:B15)", "A17": "H1 2024", "B17": "=SUM(B2:B7)",
                                         "A18": "H2 2024", "B18": "=SUM(B8:B13)"}, 14)
    assert [f.level for f in _findings(halves, before)] == ["warning", "warning"]
    # Average and Max that covered the whole year, left behind: errors that cite the original total.
    before = sales(tmp_path / "mb.xlsx", {"A14": "Total", "B14": "=SUM(B2:B13)", "A15": "Average", "B15": "=AVERAGE(B2:B13)",
                                          "A16": "Max", "B16": "=MAX(B2:B13)"})
    stale = sales(tmp_path / "m.xlsx", {"A16": "Total", "B16": "=SUM(B2:B15)", "A17": "Average", "B17": "=AVERAGE(B2:B13)",
                                        "A18": "Max", "B18": "=MAX(B2:B13)"}, 14)
    found = _findings(stale, before)
    assert [(f.level, f.where) for f in found] == [("error", "Sales!B17"), ("error", "Sales!B18")]
    assert all(f.message.endswith("(the total was in row 14: =SUM(B2:B13))") for f in found)


def test_total_left_behind_by_a_new_sum_of_the_added_rows_is_a_warning_with_the_original(tmp_path):
    rows = [["Month", "Revenue"]] + [[f"2024-{r - 1:02d}", 100 + 10 * r] for r in range(2, 14)]
    before = _book(tmp_path / "before.xlsx", rows, "Sales", {"A14": "Total", "B14": "=SUM(B2:B13)"})
    grown = rows + [["2025-01", 300], ["2025-02", 310]]
    after = _book(tmp_path / "after.xlsx", grown, "Sales", {"A16": "Total", "B16": "=SUM(B2:B13)",
                                                            "A17": "Jan-Feb 2025", "B17": "=SUM(B14:B15)"})
    found = _findings(after, before)
    assert [(f.level, f.check, f.where) for f in found] == [("warning", "range-omits-adjacent-cells", "Sales!B16")]
    assert "was the total of the whole block, rows 2-13" in found[0].message
    # Without the original the same shape is a year split (2024 / 2025): nothing to report.
    assert _findings(after) == []


def test_data_sheet_replaced_by_its_formula_texts_is_removed_not_renamed(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Month", "Revenue"])
    for i in range(12):
        ws.append([f"2024-{i + 1:02d}", 100 + 7 * i])
    ws["D1"], ws["D2"], ws["D3"] = "=SUM(B2:B13)", "=AVERAGE(B2:B13)", "=MAX(B2:B13)"
    wb.create_sheet("Report").append(["Total", "=Data!D1"])
    before = tmp_path / "before.xlsx"
    wb.save(before)

    def replace(w):  # the data is gone; a new sheet holds only the three formula texts
        del w["Data"]
        figures = w.create_sheet("Figures")
        figures["A1"], figures["A2"], figures["A3"] = "=SUM(B2:B13)", "=AVERAGE(B2:B13)", "=MAX(B2:B13)"
        w["Report"]["B1"] = "=Figures!A1"

    found = _findings(_edit(before, tmp_path / "after.xlsx", replace), before)
    assert [(f.level, f.check, f.where) for f in found] == [("error", "sheet-removed", "Data")]
    assert "renamed" not in found[0].message


def _random_ledger_edits(seed, same_day):
    """The round-4 checker's random valid edits: 1-3 inserts or deletes, the balance refilled."""
    import random

    rnd = random.Random(seed)
    items = [(f"2024-03-{((i + 1) // 2) if same_day else i:02d}", f"tx{i}", rnd.randint(-99, 99)) for i in range(1, 16)]
    new = list(items)
    for k in range(rnd.randint(1, 3)):
        if rnd.random() < 0.5 and len(new) > 3:
            del new[rnd.randint(1, len(new) - 1)]
        else:
            j = rnd.randint(1, len(new))
            new.insert(j, (new[j - 1][0], f"new{k}", rnd.randint(-99, 99)))
    return items, new


def test_running_balance_with_repeated_dates_is_not_an_error(tmp_path):
    day = [("2024-03-01", "Opening", 1000), ("2024-03-01", "Rent", -400), ("2024-03-02", "Invoice 17", 650),
           ("2024-03-02", "Fuel", -60), ("2024-03-03", "Invoice 18", 300), ("2024-03-03", "Phone", -45),
           ("2024-03-04", "Bank fee", -5), ("2024-03-04", "Invoice 19", 820), ("2024-03-05", "Coffee", -12)]
    before = _ledger(tmp_path / "day.xlsx", day)
    same_day = day[:4] + [("2024-03-02", "Refund", 25)] + day[4:]  # inserted after Fuel, same date
    assert _findings(_ledger(tmp_path / "same_day.xlsx", same_day), before) == []
    items = [(f"2024-03-{r:02d}", d, a) for r, (d, a) in enumerate(LEDGER + [("Coffee", -12), ("Invoice 20", 410)], start=2)]
    before = _ledger(tmp_path / "ledger.xlsx", items)
    mixed = items[:3] + items[4:7] + [("2024-03-08", "Refund", 15)] + items[7:]  # Fuel deleted, Refund inserted
    assert _findings(_ledger(tmp_path / "mixed.xlsx", mixed), before) == []
    for same in (False, True):
        for seed in range(40):
            old, new = _random_ledger_edits(seed, same)
            b, a = _ledger(tmp_path / f"b{seed}.xlsx", old), _ledger(tmp_path / f"a{seed}.xlsx", new)
            assert not _errors(_findings(a, b)), (seed, same)
    # Sorted by amount and refilled: right when the order was the task, so a warning, not an error.
    srt = _ledger(tmp_path / "sorted.xlsx", sorted(items, key=lambda x: x[2]))
    found = _findings(srt, before)
    assert [(f.level, f.check) for f in found] == [("warning", "reference-meaning-changed")]
    # Not refilled after an openpyxl insert: the moved balances still read the wrong rows.
    wrong = _edit(before, tmp_path / "wrong.xlsx", lambda w: w.active.insert_rows(6))
    assert [f.check for f in _errors(_findings(wrong, before))] == ["formula-moved-as-text"]


def test_whole_column_formulas_moved_by_an_inserted_row_are_not_moved_as_text(tmp_path):
    rows = [["Region", "Amount", None, "Total"]] + [[r, a] for r, a in [("N", 10), ("S", 20), ("N", 30), ("E", 40), ("S", 50)]]
    before = _book(tmp_path / "before.xlsx", rows, "S", {"E2": "=SUM(B:B)", "E3": '=SUMIF(A:A,"N",B:B)'})

    def insert(w):  # a whole row inserted, as Excel does: E2/E3 move to E3/E4 with the same text
        ws = w.active
        ws.insert_rows(2)
        ws["A2"], ws["B2"] = "W", 5

    assert _findings(_edit(before, tmp_path / "after.xlsx", insert), before) == []


def test_chart_over_a_computed_series_still_warns_when_it_stops_early(tmp_path):
    from openpyxl.chart import LineChart, Reference

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "R"
    ws.append(["M", "Cum"])
    ws.append(["M1", 5])
    for r in range(3, 12):  # each row adds the two above it: data, not a total row
        ws.append([f"M{r - 1}", f"=B{r - 1}+B{r - 2}" if r > 3 else "=B2+B2"])
    chart = LineChart()
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=6), titles_from_data=True)
    ws.add_chart(chart, "D2")
    wb.save(tmp_path / "fib.xlsx")
    found = _findings(tmp_path / "fib.xlsx")
    assert [f.level for f in found] == ["warning"] and "data continues to row 11" in found[0].message


# -- fifth review: the round-4 leniencies narrowed to the cases they were meant for -------------
MONTHS = [f"2024-{m:02d}" for m in range(1, 13)] + ["2025-01", "2025-02"]


def _sales(path, extra, n=12):
    """Months in A2:B(n+1) of a 'Sales' sheet, with ``extra`` cells (totals) below them."""
    rows = [["Month", "Revenue"]] + [[MONTHS[r - 2], 100 + 10 * r] for r in range(2, n + 2)]
    return _book(path, rows, "Sales", extra)


def test_one_off_formula_reading_an_overwritten_neighbour_row_is_an_error(tmp_path):
    # VAT =B14*0.2 under the Total: a new month written over the Total row, the Total re-added below.
    before = _sales(tmp_path / "vat_b.xlsx", {"A14": "Total", "B14": "=SUM(B2:B13)", "A15": "VAT", "B15": "=B14*0.2"})
    after = _book(tmp_path / "vat.xlsx", [["Month", "Revenue"]] + [[MONTHS[r - 2], 100 + 10 * r] for r in range(2, 14)]
                  + [["2025-01", 300]], "Sales", {"A15": "VAT", "B15": "=B14*0.2", "A16": "Total", "B16": "=SUM(B2:B14)"})
    found = _findings(after, before)
    assert [(f.level, f.check, f.where) for f in found] == [("error", "reference-meaning-changed", "Sales!B15")]
    assert "'Total' is now in row 16" in found[0].message
    # Profit =B2-B3: a 'Marketing' line written over the Cost row, Cost moved below Profit.
    before = _book(tmp_path / "pl_b.xlsx", [["Item", "Amount"], ["Revenue", 1000], ["Cost", 600], ["Profit", "=B2-B3"]], "PL")
    after = _book(tmp_path / "pl.xlsx", [["Item", "Amount"], ["Revenue", 1000], ["Marketing", 50], ["Profit", "=B2-B3"],
                                         ["Cost", 600]], "PL")
    assert [(f.level, f.where) for f in _findings(after, before)] == [("error", "PL!B4")]
    # A change column filled down keeps its leniency: a month's row corrected in place stays clean.
    rows = [["Month", "Sales", "Change"]] + [[m, 100 + 13 * i % 40, f"=B{i + 2}-B{i + 1}" if i else None]
                                            for i, m in enumerate(MONTHS[:12])]
    before = _book(tmp_path / "chg_b.xlsx", rows, "Sales")
    fixed = _edit(before, tmp_path / "chg.xlsx", lambda w: (w.active.__setitem__("A7", "2024-06 (rev.)"),
                                                            w.active.__setitem__("B7", 999)))
    assert _findings(fixed, before) == []


def test_whole_column_formulas_moved_by_a_row_and_a_column_insert_are_moved_as_text(tmp_path):
    rows = [["Region", "Amount", None, "Total"]] + [[r, a] for r, a in [("N", 10), ("S", 20), ("N", 30), ("E", 40), ("S", 50)]]
    before = _book(tmp_path / "before.xlsx", rows, "S", {"E2": "=SUM(B:B)", "E3": '=SUMIF(A:A,"N",B:B)'})

    def insert(w):  # openpyxl moves E2/E3 to F3/F4 with their texts kept; B is now the new, empty 'Code' column
        ws = w.active
        ws.insert_rows(1)
        ws["A1"] = "Sales by region"
        ws.insert_cols(2)
        ws["B2"] = "Code"

    stale = _edit(before, tmp_path / "stale.xlsx", insert)
    found = _findings(stale, before)
    assert [(f.level, f.check, f.where) for f in found] == [("error", "formula-moved-as-text", "S!F3")]
    fixed = _edit(stale, tmp_path / "fixed.xlsx", lambda w: (w.active.__setitem__("F3", "=SUM(C:C)"),
                                                             w.active.__setitem__("F4", '=SUMIF(A:A,"N",C:C)')))
    assert _findings(fixed, before) == []


def test_stale_aggregate_error_names_the_original_total_only_when_identified(tmp_path):
    # The total changed from SUBTOTAL to SUM and moved down by a spacer: identified by its 'Total' label.
    before = _sales(tmp_path / "sp_b.xlsx", {"A14": "Total", "B14": "=SUBTOTAL(9,B2:B13)", "A15": "Average",
                                             "B15": "=AVERAGE(B2:B13)", "A16": "Max", "B16": "=MAX(B2:B13)"})
    after = _sales(tmp_path / "sp.xlsx", {"A17": "Total", "B17": "=SUM(B2:B15)", "A18": "Average", "B18": "=AVERAGE(B2:B13)",
                                          "A19": "Max", "B19": "=MAX(B2:B13)"}, 14)
    found = _findings(after, before)
    assert [(f.level, f.where) for f in found] == [("error", "Sales!B18"), ("error", "Sales!B19")]
    assert all(f.message.endswith("(the total was in row 14: =SUBTOTAL(9,B2:B13))") for f in found)
    # The original had no total (Max, Average) and the Total is new: no row named, nothing 'extended'.
    before = _sales(tmp_path / "nt_b.xlsx", {"A14": "Max", "B14": "=MAX(B2:B13)", "A15": "Average", "B15": "=AVERAGE(B2:B13)"})
    after = _sales(tmp_path / "nt.xlsx", {"A16": "Total", "B16": "=SUM(B2:B15)", "A17": "Max", "B17": "=MAX(B2:B15)",
                                          "A18": "Average", "B18": "=AVERAGE(B2:B13)"}, 14)
    found = _findings(after, before)
    assert [(f.level, f.where) for f in found] == [("error", "Sales!B18")]
    assert "row 16 (=SUM(B2:B15)) covers them, but this formula does not" in found[0].message
    assert "extended" not in found[0].message and "the total was in row" not in found[0].message


def test_year_split_that_keeps_a_covering_total_is_clean_with_the_original(tmp_path):
    before = _sales(tmp_path / "before.xlsx", {"A14": "Total", "B14": "=SUM(B2:B13)"})
    parts = {"2024": "=SUM(B2:B13)", "2025": "=SUM(B14:B15)", "Total": "=SUM(B2:B15)"}
    for order in (["2024", "2025", "Total"], ["Total", "2024", "2025"]):
        extra = {}
        for r, label in enumerate(order, start=16):
            extra[f"A{r}"], extra[f"B{r}"] = label, parts[label]
        assert _findings(_sales(tmp_path / f"{order[0]}.xlsx", extra, 14), before) == [], order
    # A 'total' that misses a row covers nothing: the old total's warning stays.
    short = _sales(tmp_path / "short.xlsx", {"A16": "2024", "B16": "=SUM(B2:B13)", "A17": "2025", "B17": "=SUM(B14:B15)",
                                             "A18": "Total", "B18": "=SUM(B3:B15)"}, 14)
    assert [(f.level, f.where) for f in _findings(short, before)] == [("warning", "Sales!B16")]


def test_plus_total_under_a_numeric_header_ends_the_data_for_charts(tmp_path):
    from openpyxl.chart import LineChart, Reference

    def book(path, last):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Y"
        ws.append(["Month", 2024])  # the header of column B is a number
        for i in range(6):
            ws.append([f"M{i + 1}", 10 + i])
        ws.append(["Total", "=B2+B3+B4+B5+B6+B7"])
        chart = LineChart()
        chart.add_data(Reference(ws, min_col=2, min_row=2, max_row=last))
        ws.add_chart(chart, "D2")
        wb.save(path)
        return path

    assert _findings(book(tmp_path / "all.xlsx", 7)) == []
    found = _findings(book(tmp_path / "short.xlsx", 6))
    assert [f.level for f in found] == ["warning"] and "data continues to row 7" in found[0].message


def _named_chart_book(path, sheets, names, local=False):
    """A data sheet (M, R, C, P), defined ``names`` ((name, text), workbook-wide unless ``local``) and
    chart sheets ``(sheet, series, title)``; a series is a column number or a reference such as Data!Rev."""
    from openpyxl.chart import BarChart, Reference
    from openpyxl.workbook.defined_name import DefinedName

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["M", "R", "C", "P"])
    for i in range(1, 8):
        ws.append([f"M{i}", i * i, 7 - i, 2 * i])
    for name, text in names:
        (ws if local else wb).defined_names[name] = DefinedName(name, attr_text=text)
    for sheet, series, title in sheets:
        chart = BarChart()
        for s in series:
            if isinstance(s, str):
                chart.add_data(Reference(ws, min_col=2, min_row=2, max_row=8))
                chart.series[-1].val.numRef.f = s
            else:
                chart.add_data(Reference(ws, min_col=s, min_row=1, max_row=8), titles_from_data=True)
        if title:
            chart.title = title
        wb.create_chartsheet(sheet).add_chart(chart)
    wb.save(path)
    return path


def test_chart_sheet_over_a_defined_name_is_compared_by_the_range_it_names(tmp_path):
    rev = [("Revenue", "Data!$B$2:$B$8")]
    before = _named_chart_book(tmp_path / "b.xlsx", [("Chart1", ["Data!Revenue"], None)], rev)
    renamed = _named_chart_book(tmp_path / "renamed.xlsx", [("Revenue", ["Data!Revenue"], None)], rev)
    assert _findings(renamed, before) == []
    direct = _named_chart_book(tmp_path / "direct.xlsx", [("Revenue chart", [2], None)], rev)  # the range itself
    assert _findings(direct, before) == []
    replaced = _named_chart_book(tmp_path / "replaced.xlsx", [("Profit", [4], "Profit")], rev)
    assert [(f.level, f.check, f.where) for f in _findings(replaced, before)] == [("error", "sheet-removed", "Chart1")]
    local = [("Rev", "Data!$B$2:$B$8")]  # a name of the data sheet itself
    before = _named_chart_book(tmp_path / "lb.xlsx", [("Chart1", ["Data!Rev"], None)], local, local=True)
    replaced = _named_chart_book(tmp_path / "lr.xlsx", [("Profit", [4], None)], local, local=True)
    assert [f.level for f in _findings(replaced, before)] == ["error"]
    # A dynamic name (no plain range) pairs only the same name; another chart in its place is a warning.
    dyn = [("DynRev", "OFFSET(Data!$B$2,0,0,COUNT(Data!$B:$B),1)")]
    before = _named_chart_book(tmp_path / "db.xlsx", [("Chart1", ["Data!DynRev"], None)], dyn)
    assert _findings(_named_chart_book(tmp_path / "dr.xlsx", [("Revenue", ["Data!DynRev"], None)], dyn), before) == []
    found = _findings(_named_chart_book(tmp_path / "dx.xlsx", [("Profit", [4], "Profit")], dyn), before)
    assert [(f.level, f.check, f.where) for f in found] == [("warning", "sheet-removed", "Chart1")]
    assert "'Profit' may be it under a new name" in found[0].message
    gone = _findings(_named_chart_book(tmp_path / "dg.xlsx", [], dyn), before)
    assert [(f.level, f.check) for f in gone] == [("error", "sheet-removed")]


def test_chart_without_identity_does_not_take_another_charts_renamed_sheet(tmp_path):
    for names, ref, level in (([("Revenue", "Data!$B$2:$B$8")], "Data!Revenue", "error"),
                              ([("DynRev", "OFFSET(Data!$B$2,0,0,COUNT(Data!$B:$B),1)")], "Data!DynRev", "warning")):
        before = _named_chart_book(tmp_path / "b.xlsx", [("Chart1", [ref], None), ("Chart2", [3], None)], names)
        # Chart2 renamed 'Costs' (same series); Chart1 replaced by 'Profit'.
        after = _named_chart_book(tmp_path / "a.xlsx", [("Costs", [3], None), ("Profit", [4], None)], names)
        found = _findings(after, before)
        assert [(f.level, f.check, f.where) for f in found] == [(level, "sheet-removed", "Chart1")], ref
        assert "'Costs'" not in found[0].message


def test_cleared_ledger_row_is_an_error_and_a_cleared_date_alone_a_plain_warning(tmp_path):
    items = [(f"2024-03-{i:02d}", f"tx{i}", (-1) ** i * (10 * i + 3)) for i in range(1, 11)]
    before = _ledger(tmp_path / "before.xlsx", items)

    def clear(cols):
        def change(w):
            for col in cols:
                w.active[f"{col}7"] = None
        return change

    # A7:D7 cleared: D8 =D7+C8 now reads an empty balance, so every later balance restarts at 0.
    found = _findings(_edit(before, tmp_path / "abcd.xlsx", clear("ABCD")), before)
    assert [(f.level, f.check, f.where) for f in found] == [("error", "reference-meaning-changed", "Ledger!D8")]
    assert "D7, which it reads, held '=D6+C7' and is now empty too" in found[0].message
    # A7:C7 cleared, the balance in D7 kept: the balances stay right. A warning that claims no reordering.
    found = _findings(_edit(before, tmp_path / "abc.xlsx", clear("ABC")), before)
    assert [(f.level, f.where) for f in found] == [("warning", "Ledger!D8")]
    assert "is now empty" in found[0].message and "reordered" not in found[0].message


def test_no_recalc_help_says_error_values_are_not_checked(capsys):
    from failbook.cli import main as failbook_main

    for run in (lambda: failbook_main(["audit", "--help"]), lambda: main(["--help"])):
        with pytest.raises(SystemExit) as exit_info:
            run()
        assert exit_info.value.code == 0
        text = " ".join(capsys.readouterr().out.split())
        assert audit_mod.NO_RECALC_HELP in text and "saved in the file" not in text
        assert "3 a workbook could not be read" in text


def test_stray_formatted_cell_far_down_does_not_inflate_the_work(tmp_path):
    from openpyxl.styles import Font

    rows = [["Item", "Qty", "Price", "Total"]] + [[f"I{r}", r, 1.5, f"=B{r}*C{r}"] for r in range(2, 22)]
    before = _book(tmp_path / "before.xlsx", rows, "Data", {"A22": "Total", "D22": "=SUM(D2:D21)"})
    wb = openpyxl.load_workbook(before)
    wb.active["H1048576"].font = Font(bold=True)  # styled, empty: the used range now ends at row 1048576
    wb.save(before)
    after = _edit(before, tmp_path / "after.xlsx", lambda w: w.active.__setitem__("B3", 7))
    old, new = openpyxl.load_workbook(before), openpyxl.load_workbook(after)
    stored = len(new.active._cells), len(old.active._cells)
    found = audit_mod._static_checks(new, after)
    audit_mod.compare(old, new, found)
    check_errors(new, new, found)
    assert not _errors(found)
    # The checks visited only stored cells; iter_rows() would have created about 8.4 million.
    assert (len(new.active._cells), len(old.active._cells)) == stored


def test_unreadable_workbook_is_a_message_and_exit_code_3(tmp_path, capsys):
    good = load_case("xlsx-001").input_dir / "sales.xlsx"
    truncated = tmp_path / "truncated.xlsx"
    truncated.write_bytes(good.read_bytes()[:3000])
    encrypted = tmp_path / "encrypted.xlsx"
    encrypted.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + bytes(4000))  # the OLE container of an encrypted file
    cases = [([str(truncated)], "truncated"), ([str(encrypted)], "password-protected"),
             ([str(tmp_path / "missing.xlsx")], "No such file"), ([str(good), "--before", str(truncated)], "--before")]
    for argv, reason in cases:
        assert main(argv + ["--no-recalc"]) == 3
        captured = capsys.readouterr()
        assert captured.out == "" and "Traceback" not in captured.err and reason in captured.err


def test_fallback_chart_reader_takes_the_title_not_the_axis_titles():
    from failbook.xlsx import charts

    path = load_case("xlsx-011").failure_dir / "monthly.xlsx"
    full = [c["title"] for c in charts(path)]
    assert full and all(full)
    assert [c["title"] for c in audit_mod._charts_fallback(path)] == full


# -- error values --------------------------------------------------------------------------------
def test_name_error_from_newer_function_is_a_warning_but_a_broken_name_is_an_error():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A2"] = 3
    ws["B2"] = "=_xlfn.LET(_xlpm.x,A2,_xlpm.x*2)"      # correct in Excel 365; #NAME? where LET is missing
    ws["B3"] = "=B2+1"                                  # inherits that #NAME?
    ws["B4"] = "=SUM(NoSuchName)"                       # a real broken name
    ws["B5"] = "=_xlfn.XLOOKUP(A2,Gone!A:A,Gone!B:B)"  # newer function, but its sheet is missing
    values = openpyxl.Workbook()
    vs = values.active
    vs.title = "S"
    for cell in ("B2", "B3", "B4", "B5"):
        vs[cell] = "#NAME?"
    found = []
    check_errors(wb, values, found)
    level = {f.where: f.level for f in found}
    assert level == {"S!B2": "warning", "S!B3": "warning", "S!B4": "error", "S!B5": "error"}


def test_missing_libreoffice_is_reported_not_raised(tmp_path, monkeypatch):
    case = load_case("xlsx-006")
    name = case.outputs[0]
    monkeypatch.setenv("FAILBOOK_SOFFICE", str(tmp_path / "no-such-soffice"))
    found = audit(case.failure_dir / name, case.input_dir / name)
    unavailable = [f for f in found if f.check == "recalculation-unavailable"]
    assert [f.level for f in unavailable] == ["warning"] and "could not be started" in unavailable[0].message
    assert "error" in _levels(found, "stale-sheet-reference")  # the other checks still ran


@pytest.mark.skipif(os.name == "nt", reason="uses a POSIX shell script as a fake LibreOffice")
def test_failing_or_hanging_libreoffice_is_reported_not_raised(tmp_path, monkeypatch):
    case = load_case("xlsx-002")
    path = case.input_dir / case.outputs[0]
    failing = tmp_path / "failing.sh"
    failing.write_text("#!/bin/sh\necho 'cannot open display' >&2\nexit 3\n")
    hanging = tmp_path / "hanging.sh"
    hanging.write_text("#!/bin/sh\nsleep 30 &\nsleep 30\n")
    for script in (failing, hanging):
        script.chmod(0o755)
    monkeypatch.setattr(audit_mod, "RECALC_TIMEOUT", 1)
    monkeypatch.setenv("FAILBOOK_SOFFICE", str(failing))
    msg = [f.message for f in audit(path) if f.check == "recalculation-unavailable"]
    assert msg and "exit code 3" in msg[0] and "cannot open display" in msg[0]
    monkeypatch.setenv("FAILBOOK_SOFFICE", str(hanging))
    msg = [f.message for f in audit(path) if f.check == "recalculation-unavailable"]
    assert msg and "did not finish within 1 s" in msg[0]


def test_skipped_recalculation_is_stated():
    case = load_case("xlsx-002")
    found = audit(case.input_dir / case.outputs[0], recalc=False)
    assert "recalculation-skipped" in {f.check for f in found}


def test_failbook_audit_forwards_the_recalculation_options(capsys):
    from failbook.cli import main as failbook_main

    case = load_case("xlsx-002")
    failbook_main(["audit", str(case.input_dir / case.outputs[0]), "--no-recalc", "--json"])
    assert "recalculation-skipped" in capsys.readouterr().out


# -- forced recalculation (claims F9) ------------------------------------------------------------
def _stale_copy(tmp_path):
    """xlsx-001's input with Qty in D2 changed in the XML only, so every cached result is stale."""
    src = load_case("xlsx-001").input_dir / "sales.xlsx"
    cached = openpyxl.load_workbook(src, data_only=True).active
    if cached["F2"].value is None:
        pytest.skip("the committed input has no cached results")
    edited = tmp_path / "data.xlsx"
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(edited, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            data = zi.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                text, n = re.subn(r'(<c r="D2"[^>]*>)<v>[^<]*</v>', r"\g<1><v>60</v>", data.decode(), count=1)
                if n != 1:
                    pytest.skip("D2 is not a plain number in the input's XML")
                data = text.encode()
            zo.writestr(item, data)
    return edited, cached["E2"].value * 60, cached["F2"].value


@needs_libreoffice
def test_recalculated_copy_recomputes_stale_cached_values(tmp_path, capsys):
    edited, fresh, stale = _stale_copy(tmp_path)
    copy = tmp_path / "recalc" / "data.xlsx"
    main([str(edited), "--recalculated", str(copy)])
    assert f"recalculated copy: {copy}" in capsys.readouterr().out
    value = openpyxl.load_workbook(copy, data_only=True).active["F2"].value
    assert value == pytest.approx(fresh) and value != pytest.approx(stale)


@needs_libreoffice
def test_skill_manual_recalculation_command_recomputes_stale_values(tmp_path):
    """The fallback command SKILL.md gives must really recalculate, unlike a plain --convert-to."""
    edited, fresh, stale = _stale_copy(tmp_path)
    text = (SKILL / "skill" / "SKILL.md").read_text(encoding="utf-8")
    block = next(b for b in re.findall(r"```bash\n(.*?)```", text, re.S) if "OOXMLRecalcMode" in b)
    script = "\n".join(line.strip() for line in block.splitlines()).replace("/tmp/", f"{tmp_path}/tmp/")
    (tmp_path / "tmp").mkdir()
    soffice = os.environ.get("FAILBOOK_SOFFICE") or shutil.which("soffice") or shutil.which("libreoffice")
    script = re.sub(r"^soffice ", f"'{soffice}' ", script, flags=re.M)
    subprocess.run(["sh", "-c", script], cwd=tmp_path, check=True, capture_output=True, timeout=300)
    value = openpyxl.load_workbook(tmp_path / "tmp" / "recalc" / "data.xlsx", data_only=True).active["F2"].value
    assert value == pytest.approx(fresh) and value != pytest.approx(stale)


# -- installing the skill fix ----------------------------------------------------------------------
def _fix_copy(tmp_path):
    src = tmp_path / "fix-src"
    shutil.copytree(SKILL, src, ignore=shutil.ignore_patterns("__pycache__"))
    pyc = src / "skill" / "scripts" / "__pycache__"
    pyc.mkdir()
    (pyc / "audit_xlsx.cpython-313.pyc").write_bytes(b"local build product")
    return load_fix(str(src))


@pytest.mark.parametrize("agent", ["codex", "gemini-cli"])
def test_agents_without_skill_discovery_get_the_scripts(tmp_path, agent):
    fix = _fix_copy(tmp_path)
    preset = PRESETS[agent]
    work = tmp_path / "work"
    work.mkdir()
    assert fix.apply(work, "do the task", preset) == "do the task"
    instructions = (work / preset.instructions_file).read_text(encoding="utf-8")
    assert "<this-skill-dir>" not in instructions
    commands = re.findall(r"python3 (\S+audit_xlsx\.py)", instructions)
    assert commands and all((work / c).is_file() for c in commands)
    assert (work / commands[0]).read_bytes() == (ROOT / "src" / "failbook" / "audit.py").read_bytes()
    assert not [p for p in work.rglob("*") if p.name == "__pycache__" or p.suffix == ".pyc"]


def test_skill_discovery_copy_leaves_out_local_build_products(tmp_path):
    fix = _fix_copy(tmp_path)
    preset = PRESETS["claude-code"]
    work = tmp_path / "work"
    work.mkdir()
    fix.apply(work, "do the task", preset)
    installed = work / preset.skills_dir / fix.skill_name
    assert (installed / "SKILL.md").is_file() and (installed / "scripts" / "audit_xlsx.py").is_file()
    assert not [p for p in work.rglob("*") if p.name == "__pycache__" or p.suffix == ".pyc"]


def test_skill_script_is_the_package_audit():
    shipped = SKILL / "skill" / "scripts" / "audit_xlsx.py"
    assert shipped.read_bytes() == (ROOT / "src" / "failbook" / "audit.py").read_bytes()
    proc = subprocess.run([sys.executable, "-I", str(shipped), "--help"], capture_output=True, text=True, timeout=60)
    assert proc.returncode == 0 and "--recalculated" in proc.stdout
