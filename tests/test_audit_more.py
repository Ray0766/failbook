"""The audit shipped with spreadsheet-verify-skill: no errors on ordinary correct edits, errors kept
on real breakage, recalculation that is forced and fails softly, and the skill installed whole."""

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
