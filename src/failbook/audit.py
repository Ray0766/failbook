#!/usr/bin/env python3
"""Audit an edited .xlsx for the silent mistakes behind Failbook's spreadsheet cases.

    python3 audit_xlsx.py AFTER.xlsx [--before BEFORE.xlsx] [--json]

It knows nothing about the task: it looks for symptoms that are almost always
mistakes (error values after recalculation, totals whose range stops short of
the data, formulas moved as text, approximate lookups, sums that include
subtotals, charts/validation/highlighting that stop before the data ends,
numbers stored as text) and, given the original file, for collateral damage
(lost sheets, formulas replaced by constants, lost validation, styles or frozen
panes, formulas whose referenced columns changed meaning).

Findings are leads, not verdicts: "error" means very likely wrong, "warning"
means check it. Exit code 1 when there is at least one error.

This file is self-contained (Python 3.10+, openpyxl; LibreOffice optional for
recalculation) so that it can be shipped inside an Agent Skill. It is the
canonical copy; fixes/spreadsheet-verify-skill/skill/scripts/audit_xlsx.py is
generated from it by `failbook docs`.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import asdict, dataclass
from pathlib import Path

import openpyxl
from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.utils.cell import column_index_from_string, get_column_letter

ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}
AGGREGATES = {"SUM", "AVERAGE", "COUNT", "COUNTA", "MIN", "MAX", "SUBTOTAL", "PRODUCT"}
SUMMARY_WORDS = re.compile(r"summary|total|dashboard|report|汇总|合计", re.I)
NUMBER_TEXT = re.compile(r"^\s*[-+(]?\s*[$€£¥]?\s*\d{1,3}(,\d{3})+(\.\d+)?\s*\)?\s*$|^\s*[-+]?[$€£¥]?\d+\.\d+\s*%?\s*$|^\s*[-+]?\d+\s*%\s*$")
DATE_TEXT = re.compile(r"^\s*\d{4}-\d{2}-\d{2}\s*$")
REF = re.compile(
    r"^(?:(?:'(?P<qs>(?:[^']|'')+)'|(?P<s>[^'!:]+))!)?"
    r"(?:\$?(?P<c1>[A-Za-z]{1,3})\$?(?P<r1>\d+)(?::\$?(?P<c2>[A-Za-z]{1,3})\$?(?P<r2>\d+))?"
    r"|\$?(?P<cc1>[A-Za-z]{1,3}):\$?(?P<cc2>[A-Za-z]{1,3}))$"
)


@dataclass
class Finding:
    level: str  # error | warning
    check: str
    where: str
    message: str


# -- helpers ----------------------------------------------------------------------------
def _formula(value) -> str | None:
    text = getattr(value, "text", value)  # ArrayFormula
    if isinstance(text, str) and text.startswith("=") and len(text) > 1:
        return text
    return None


def _ref(sheet: str, row: int, col: int) -> str:
    q = sheet if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", sheet) else "'" + sheet.replace("'", "''") + "'"
    return f"{q}!{get_column_letter(col)}{row}"


def parse_range(text: str, default_sheet: str):
    """'Sheet'!$B$2:$B$13 -> (sheet, c1, r1, c2, r2); whole columns get r1=1, r2=None."""
    text = text.strip().lstrip("=")
    if text.startswith("(") and text.endswith(")"):
        text = text[1:-1]
    m = REF.match(text)
    if not m:
        return None
    sheet = (m.group("qs") or "").replace("''", "'") or m.group("s") or default_sheet
    if m.group("cc1"):
        c1, c2 = column_index_from_string(m.group("cc1").upper()), column_index_from_string(m.group("cc2").upper())
        return sheet, min(c1, c2), 1, max(c1, c2), None
    c1, r1 = column_index_from_string(m.group("c1").upper()), int(m.group("r1"))
    c2 = column_index_from_string(m.group("c2").upper()) if m.group("c2") else c1
    r2 = int(m.group("r2")) if m.group("r2") else r1
    return sheet, min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2)


def _tokens(formula: str):
    try:
        return Tokenizer(formula).items
    except Exception:
        return []


def _refs(formula: str, sheet: str):
    for tok in _tokens(formula):
        if tok.type == Token.OPERAND and tok.subtype == Token.RANGE:
            rng = parse_range(tok.value, sheet)
            if rng:
                yield tok.value, rng


def _function_calls(formula: str):
    """Yield (FUNCTION, [argument texts]) for every call, nested ones included."""
    toks = _tokens(formula)
    stack = []
    for tok in toks:
        if tok.type == Token.FUNC and tok.subtype == Token.OPEN:
            stack.append([tok.value[:-1].upper().replace("_XLFN.", ""), [""]])
        elif tok.type == Token.FUNC and tok.subtype == Token.CLOSE and stack:
            name, args = stack.pop()
            if stack:
                stack[-1][1][-1] += f"{name}(...)"
            yield name, [a for a in args if a != ""] if args != [""] else []
        elif tok.type == Token.SEP and tok.subtype == Token.ARG and stack:
            stack[-1][1].append("")
        elif stack:
            stack[-1][1][-1] += tok.value


def _has_value(ws, row: int, col: int) -> bool:
    v = ws.cell(row, col).value
    return v is not None and not (isinstance(v, str) and not v.strip())


def _last_filled(ws, col: int, start: int) -> int:
    """Last row of the contiguous run of filled cells in ``col`` starting at ``start``."""
    r = start
    while _has_value(ws, r + 1, col):
        r += 1
    return r


def _cells(sqref) -> list[tuple[int, int, int, int]]:
    out = []
    for rng in str(sqref).split():
        m = re.match(r"^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$", rng)
        if m:
            c1, r1 = column_index_from_string(m.group(1)), int(m.group(2))
            c2 = column_index_from_string(m.group(3)) if m.group(3) else c1
            r2 = int(m.group(4)) if m.group(4) else r1
            out.append((c1, r1, c2, r2))
    return out


# -- recalculation ----------------------------------------------------------------------
_XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
"""


def recalculated(path: Path):
    """Workbook (data_only) recalculated by LibreOffice, or None if LibreOffice is unavailable."""
    soffice = os.environ.get("FAILBOOK_SOFFICE") or shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    with tempfile.TemporaryDirectory(prefix="audit-") as tmp:
        tmp = Path(tmp)
        (tmp / "profile" / "user").mkdir(parents=True)
        (tmp / "profile" / "user" / "registrymodifications.xcu").write_text(_XCU, encoding="utf-8")
        src = tmp / "in.xlsx"
        shutil.copyfile(path, src)
        subprocess.run(
            [soffice, "--headless", "--norestore", "--nolockcheck", f"-env:UserInstallation={(tmp / 'profile').as_uri()}",
             "--convert-to", "xlsx:Calc Office Open XML", "--outdir", str(tmp / "out"), str(src)],
            capture_output=True, timeout=180,
        )
        out = tmp / "out" / "in.xlsx"
        return openpyxl.load_workbook(out, data_only=True) if out.exists() else None


# -- checks on the edited file ------------------------------------------------------------
def check_errors(wb, values, findings):
    if values is None:
        findings.append(Finding("warning", "recalculation-unavailable", "-",
                                "LibreOffice not found; error values could not be checked. Install it or open the file in a spreadsheet app."))
        return
    for ws in values.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.strip().upper() in ERRORS:
                    f = _formula(wb[ws.title][c.coordinate].value) if ws.title in wb.sheetnames else None
                    findings.append(Finding("error", "error-value", _ref(ws.title, c.row, c.column),
                                            f"evaluates to {c.value}" + (f" ({f})" if f else "")))


def check_formulas(wb, findings):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                f = _formula(c.value)
                if not f:
                    continue
                where = _ref(ws.title, c.row, c.column)
                for name, args in _function_calls(f):
                    if name in ("VLOOKUP", "HLOOKUP"):
                        last = args[3].strip().upper() if len(args) >= 4 else None
                        if last in (None, "TRUE", "1"):
                            findings.append(Finding("warning", "approximate-lookup", where,
                                                    f"{name} without exact match (4th argument FALSE): {f}"))
                    if name == "MATCH":
                        last = args[2].strip() if len(args) >= 3 else None
                        if last in (None, "1", "-1"):
                            findings.append(Finding("warning", "approximate-lookup", where, f"MATCH without match_type 0: {f}"))
                for text, (sheet, c1, r1, c2, r2) in _refs(f, ws.title):
                    if sheet != ws.title or r2 is None or c1 != c2 or r2 - r1 < 1:
                        continue
                    head = f[1:].lstrip("+").upper()
                    agg = next((a for a in AGGREGATES if head.startswith(a + "(")), None)
                    if agg is None:
                        continue
                    # Total below a column range that stops short of the data above it.
                    if c1 == c.column and r2 < c.row - 1:
                        gap = [r for r in range(r2 + 1, c.row) if _has_value(ws, r, c1)]
                        if gap:
                            findings.append(Finding("error", "range-omits-adjacent-cells", where,
                                                    f"{text} stops at row {r2}, but rows {gap[0]}-{gap[-1]} between it and the formula hold values"))
                    # A sum over a range that contains subtotals of the same range.
                    if agg in ("SUM", "PRODUCT"):
                        inner = []
                        for r in range(r1, r2 + 1):
                            g = _formula(ws.cell(r, c1).value)
                            if g and g[1:].upper().startswith("SUM(") and any(
                                s == ws.title and cc1 == c1 and rr2 is not None and r1 <= rr1 and rr2 <= r2
                                for _, (s, cc1, rr1, _cc2, rr2) in _refs(g, ws.title)
                            ):
                                inner.append(r)
                        if inner:
                            findings.append(Finding("error", "double-counting", where,
                                                    f"{text} includes rows {', '.join(map(str, inner[:6]))} which are themselves sums of rows in the same range"))


def check_ranges(wb, path, findings):
    """Charts, data validation and conditional formatting that stop before the data ends."""
    def short(ws, c1, r1, c2, r2, what, where):
        for col in range(c1, c2 + 1):
            if not _has_value(ws, r2, col):
                continue
            end = _last_filled(ws, col, r2)
            if end > r2:
                findings.append(Finding("warning", "range-stops-before-data-ends", where,
                                        f"{what} covers rows {r1}-{r2} of column {get_column_letter(col)}, but data continues to row {end}"))
                return

    for ws in wb.worksheets:
        for dv in ws.data_validations.dataValidation:
            for c1, r1, c2, r2 in _cells(dv.sqref):
                short(ws, c1, r1, c2, r2, f"data validation ({dv.type})", f"{ws.title}!{dv.sqref}")
        for cf in ws.conditional_formatting:
            for c1, r1, c2, r2 in _cells(cf.sqref):
                short(ws, c1, r1, c2, r2, "conditional formatting", f"{ws.title}!{cf.sqref}")
    try:
        from failbook.xlsx import charts as _charts  # the package's fuller chart reader, when installed
    except Exception:
        _charts = _charts_fallback
    for chart in _charts(path):
        for s in chart["series"]:
            rng = parse_range(s.get("val") or "", chart["sheet"])
            if rng and rng[0] in wb.sheetnames and rng[4]:
                sheet, c1, r1, c2, r2 = rng
                short(wb[sheet], c1, r1, c2, r2, f"chart '{chart['title']}' series", f"{chart['sheet']} (chart)")


def _charts_fallback(path):
    """Minimal chart reader used when the failbook package is not importable."""
    import posixpath
    import zipfile
    from xml.etree import ElementTree as ET

    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "p": "http://schemas.openxmlformats.org/package/2006/relationships",
          "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
          "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

    def rels(z, part):
        rp = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
        if rp not in z.namelist():
            return {}
        out = {}
        for rel in ET.fromstring(z.read(rp)).findall("p:Relationship", ns):
            t = rel.get("Target", "")
            out[rel.get("Id")] = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(part), t))
        return out

    found = []
    with zipfile.ZipFile(path) as z:
        wbrels = rels(z, "xl/workbook.xml")
        for sh in ET.fromstring(z.read("xl/workbook.xml")).findall("m:sheets/m:sheet", ns):
            part = wbrels.get(sh.get(f"{{{ns['r']}}}id"))
            if not part or part not in z.namelist():
                continue
            srels = rels(z, part)
            for d in ET.fromstring(z.read(part)).findall("m:drawing", ns):
                dpart = srels.get(d.get(f"{{{ns['r']}}}id"))
                if not dpart or dpart not in z.namelist():
                    continue
                for cpart in rels(z, dpart).values():
                    if "/charts/" not in cpart or cpart not in z.namelist():
                        continue
                    root = ET.fromstring(z.read(cpart))
                    title = "".join(t.text or "" for t in root.iter(f"{{{ns['a']}}}t"))
                    series = [{"val": (ser.find("c:val/c:numRef/c:f", ns).text if ser.find("c:val/c:numRef/c:f", ns) is not None else None)}
                              for ser in root.iter(f"{{{ns['c']}}}ser")]
                    found.append({"sheet": sh.get("name"), "title": title, "series": series})
    return found


def check_text_numbers(wb, findings):
    for ws in wb.worksheets:
        cols: dict[int, list] = {}
        for row in ws.iter_rows(min_row=2):
            for c in row:
                if c.value is not None:
                    cols.setdefault(c.column, []).append(c)
        for col, cells in cols.items():
            numeric = sum(1 for c in cells if isinstance(c.value, (int, float, dt.date, dt.datetime)) and not isinstance(c.value, bool))
            texty = [c for c in cells if isinstance(c.value, str) and (NUMBER_TEXT.match(c.value) or DATE_TEXT.match(c.value))]
            if texty and numeric >= len(texty):
                first = texty[0]
                findings.append(Finding("error", "number-stored-as-text", _ref(ws.title, first.row, col),
                                        f"{len(texty)} cell(s) in column {get_column_letter(col)} hold text that looks like a number or date "
                                        f"(e.g. {first.value!r}); sums and date filters ignore them"))


def check_pasted_totals(wb, findings):
    for ws in wb.worksheets:
        formulas = numbers = 0
        labels = []
        for row in ws.iter_rows():
            for c in row:
                if _formula(c.value):
                    formulas += 1
                elif isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                    numbers += 1
                elif isinstance(c.value, str):
                    labels.append(c.value)
        if formulas == 0 and numbers >= 3 and (SUMMARY_WORDS.search(ws.title) or any(SUMMARY_WORDS.search(t) for t in labels[:50])):
            findings.append(Finding("warning", "totals-without-formulas", ws.title,
                                    f"'{ws.title}' holds {numbers} numbers and no formulas; if they summarise other sheets they will not update"))


# -- comparison with the original ----------------------------------------------------------
def compare(before, after, findings):
    def content(ws):
        return {c.coordinate: c.value for row in ws.iter_rows() for c in row if c.value is not None}

    added = [n for n in after.sheetnames if n not in before.sheetnames]
    for name in before.sheetnames:
        if name in after.sheetnames:
            continue
        old = content(before[name])
        renamed_to = next((n for n in added if old and content(after[n]) == old), None)
        if renamed_to is None:
            findings.append(Finding("error", "sheet-removed", name, f"sheet '{name}' existed before and is gone"))
            continue
        # A rename: fine, as long as nothing still points at the old name.
        pattern = re.compile(r"(?<![A-Za-z0-9_.])(?:'" + re.escape(name.replace("'", "''")) + "'|" + re.escape(name) + r")!")
        stale = [_ref(ws.title, c.row, c.column) for ws in after.worksheets for row in ws.iter_rows() for c in row
                 if _formula(c.value) and pattern.search(_formula(c.value))]
        stale += [f"name {dn}" for dn, d in after.defined_names.items() if pattern.search(str(d.attr_text or ""))]
        if stale:
            findings.append(Finding("error", "stale-sheet-reference", stale[0],
                                    f"sheet '{name}' was renamed to '{renamed_to}', but {len(stale)} formula(s)/name(s) still refer to '{name}'"))
    for name in set(before.sheetnames) & set(after.sheetnames):
        b, a = before[name], after[name]
        fb = sum(1 for row in b.iter_rows() for c in row if _formula(c.value))
        fa = sum(1 for row in a.iter_rows() for c in row if _formula(c.value))
        if fb >= 3 and fa < fb / 2:
            findings.append(Finding("error", "formulas-lost", name, f"'{name}' had {fb} formulas and now has {fa}"))
        if len(b.data_validations.dataValidation) > len(a.data_validations.dataValidation):
            findings.append(Finding("error", "validation-lost", name, "data validation rules were removed"))
        if b.freeze_panes and not a.freeze_panes:
            findings.append(Finding("warning", "frozen-panes-lost", name, f"frozen panes at {b.freeze_panes} were removed"))
        heads_b = {str(c.value): c for c in b[1] if c.value is not None}
        for c in a[1]:
            old = heads_b.get(str(c.value))
            if old is not None and bool(old.font and old.font.b) and not bool(c.font and c.font.b):
                findings.append(Finding("warning", "header-style-lost", _ref(name, 1, c.column), f"header '{c.value}' lost its bold style"))
                break
        # Formula text that moved to another row or column: references were not updated.
        positions_b: dict[str, list[tuple[int, int]]] = {}
        for row in b.iter_rows():
            for c in row:
                f = _formula(c.value)
                if f:
                    positions_b.setdefault(f, []).append((c.row, c.column))
        moved = []
        for row in a.iter_rows():
            for c in row:
                f = _formula(c.value)
                if not f or "$" in f or f not in positions_b or (c.row, c.column) in positions_b[f]:
                    continue
                if any(rng[0] == name for _, rng in _refs(f, name)):
                    moved.append((c, positions_b[f][0]))
        if moved:
            c, (r0, c0) = moved[0]
            findings.append(Finding("error", "formula-moved-as-text", _ref(name, c.row, c.column),
                                    f"{len(moved)} formula(s) were moved without updating their references, e.g. {c.value} was at "
                                    f"{get_column_letter(c0)}{r0}"))
    # Unchanged formulas whose referenced columns or rows now mean something else.
    def label(wb, sheet, row, col, axis):
        if sheet not in wb.sheetnames:
            return None
        ws = wb[sheet]
        return ws.cell(1, col).value if axis == "col" else ws.cell(row, 1).value

    for name in set(before.sheetnames) & set(after.sheetnames):
        for row in after[name].iter_rows():
            for c in row:
                f = _formula(c.value)
                if not f or _formula(before[name][c.coordinate].value) != f:
                    continue
                for text, (sheet, c1, r1, c2, r2) in _refs(f, name):
                    if sheet not in after.sheetnames or sheet not in before.sheetnames:
                        continue
                    if c1 == c2 and (sheet != name or r2 is None or r2 - r1 >= 1):
                        hb, ha = label(before, sheet, r1, c1, "col"), label(after, sheet, r1, c1, "col")
                        if hb != ha and r1 > 1:
                            findings.append(Finding("error", "reference-meaning-changed", _ref(name, c.row, c.column),
                                                    f"{f} refers to column {get_column_letter(c1)} of '{sheet}', whose header was {hb!r} and is now {ha!r}"))
                            break
                    if r2 == r1 and not (sheet == name and r1 == c.row):
                        lb, la = label(before, sheet, r1, c1, "row"), label(after, sheet, r1, c1, "row")
                        if lb != la:
                            findings.append(Finding("error", "reference-meaning-changed", _ref(name, c.row, c.column),
                                                    f"{f} refers to row {r1} of '{sheet}', which was {lb!r} and is now {la!r}"))
                            break


# -- entry points ---------------------------------------------------------------------------
def audit(after: Path, before: Path | None = None, recalc: bool = True) -> list[Finding]:
    findings: list[Finding] = []
    wb = openpyxl.load_workbook(after)
    if recalc:
        check_errors(wb, recalculated(Path(after)), findings)
    check_formulas(wb, findings)
    check_ranges(wb, Path(after), findings)
    check_text_numbers(wb, findings)
    check_pasted_totals(wb, findings)
    if before is not None:
        compare(openpyxl.load_workbook(before), wb, findings)
    return findings


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    p.add_argument("after", help="the edited workbook")
    p.add_argument("--before", help="the original workbook (enables collateral-damage checks)")
    p.add_argument("--no-recalc", action="store_true", help="skip LibreOffice recalculation")
    p.add_argument("--json", action="store_true")
    args = p.parse_args(argv)
    findings = audit(Path(args.after), Path(args.before) if args.before else None, recalc=not args.no_recalc)
    if args.json:
        print(json.dumps([asdict(f) for f in findings], ensure_ascii=False, indent=2))
    else:
        if not findings:
            print("no findings")
        for f in findings:
            print(f"[{f.level}] {f.check} at {f.where}: {f.message}")
    return 1 if any(f.level == "error" for f in findings) else 0


if __name__ == "__main__":
    sys.exit(main())
