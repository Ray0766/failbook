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
    assert "in the original file both ended at row 13" in found[0].message
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
