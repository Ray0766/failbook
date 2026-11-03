#!/usr/bin/env python3
"""Audit an edited .xlsx for the silent mistakes behind Failbook's spreadsheet cases.

    python3 audit_xlsx.py AFTER.xlsx [--before BEFORE.xlsx] [--recalculated COPY.xlsx] [--json]

It knows nothing about the task: it looks for symptoms that are almost always
mistakes (error values after recalculation, totals whose range stops short of
the data, formulas moved as text, approximate lookups, sums that include
subtotals, charts/validation/highlighting that stop before the data ends,
numbers stored as text) and, given the original file, for collateral damage
(lost sheets, formulas replaced by constants, lost validation, styles or frozen
panes, formulas whose referenced columns changed meaning).

Findings are leads, not verdicts: "error" means very likely wrong, "warning"
means check it. Exit code 1 when there is at least one error. A symptom that
the original file (--before) already shows is reported as a warning marked
"also in the original file": the edit did not cause it, but the task may still
need it changed.

Recalculation forces LibreOffice to recompute every formula (a plain
`soffice --convert-to xlsx` keeps the results already cached in the file).
--recalculated saves that recalculated copy, to read with
openpyxl.load_workbook(COPY, data_only=True).

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
import signal
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


def _value(ws, row: int, col: int):
    """A cell's value without creating the cell (``ws.cell`` would grow the sheet's used range)."""
    cells = getattr(ws, "_cells", None)
    if cells is None:
        return ws.cell(row, col).value
    c = cells.get((row, col))
    return None if c is None else c.value


def _has_value(ws, row: int, col: int) -> bool:
    v = _value(ws, row, col)
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


#: Seconds LibreOffice may take to recalculate one workbook before it is stopped.
RECALC_TIMEOUT = 180


class RecalcUnavailable(Exception):
    """LibreOffice could not recalculate the workbook; the message says why."""


def _stop(proc) -> None:
    """Kill LibreOffice and the processes it started (it runs in its own session)."""
    try:
        if hasattr(os, "killpg"):
            os.killpg(proc.pid, signal.SIGKILL)
        else:
            proc.kill()
    except OSError:
        pass
    try:
        proc.wait(timeout=10)
    except subprocess.TimeoutExpired:
        pass


def recalculate(path: Path, dest: Path, timeout: int | None = None) -> Path:
    """Write a copy of ``path`` to ``dest`` in which LibreOffice recalculated every formula.

    The temporary profile sets OOXMLRecalcMode=0 ("always recalculate"); with the
    default setting LibreOffice keeps the results cached in the file. Raises
    RecalcUnavailable, naming the cause, when LibreOffice is missing, fails or hangs.
    """
    timeout = RECALC_TIMEOUT if timeout is None else timeout
    soffice = os.environ.get("FAILBOOK_SOFFICE") or shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RecalcUnavailable("LibreOffice was not found (no soffice or libreoffice on PATH)")
    with tempfile.TemporaryDirectory(prefix="audit-", ignore_cleanup_errors=True) as tmp:
        tmp = Path(tmp)
        (tmp / "profile" / "user").mkdir(parents=True)
        (tmp / "profile" / "user" / "registrymodifications.xcu").write_text(_XCU, encoding="utf-8")
        src = tmp / "in.xlsx"
        shutil.copyfile(path, src)
        cmd = [soffice, "--headless", "--norestore", "--nolockcheck", f"-env:UserInstallation={(tmp / 'profile').as_uri()}",
               "--convert-to", "xlsx:Calc Office Open XML", "--outdir", str(tmp / "out"), str(src)]
        with open(tmp / "soffice.log", "wb") as log:
            try:
                proc = subprocess.Popen(cmd, stdout=log, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
                                        start_new_session=True)
            except OSError as exc:
                raise RecalcUnavailable(f"LibreOffice could not be started ({soffice}: {exc.strerror or exc})") from None
            try:
                code = proc.wait(timeout=timeout)
            except subprocess.TimeoutExpired:
                _stop(proc)
                raise RecalcUnavailable(f"LibreOffice ({soffice}) did not finish within {timeout} s and was stopped") from None
        out = tmp / "out" / "in.xlsx"
        if not out.is_file():
            lines = (tmp / "soffice.log").read_text(encoding="utf-8", errors="replace").strip().splitlines()
            detail = f": {lines[-1][:200]}" if lines else ""
            raise RecalcUnavailable(f"LibreOffice ({soffice}) wrote no recalculated file (exit code {code}{detail})")
        dest = Path(dest)
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(out, dest)
    return dest


def recalculated(path: Path, keep: Path | None = None):
    """Workbook (data_only) recalculated by LibreOffice; ``keep`` also saves the recalculated copy.

    Raises RecalcUnavailable when LibreOffice cannot produce it.
    """
    with tempfile.TemporaryDirectory(prefix="audit-out-", ignore_cleanup_errors=True) as tmp:
        dest = recalculate(Path(path), Path(keep) if keep else Path(tmp) / "recalculated.xlsx")
        try:
            return openpyxl.load_workbook(dest, data_only=True)
        except Exception as exc:  # a truncated or invalid file from LibreOffice
            raise RecalcUnavailable(f"the file LibreOffice wrote could not be read ({type(exc).__name__}: {exc})") from None


# -- checks on the edited file ------------------------------------------------------------
#: Excel 2019/2021/365 functions that LibreOffice added late or not at all; on an
#: engine without them a correct formula evaluates to #NAME?.
NEWER_FUNCTIONS = {
    "LET", "LAMBDA", "MAP", "REDUCE", "SCAN", "MAKEARRAY", "BYROW", "BYCOL", "ISOMITTED", "XLOOKUP", "XMATCH",
    "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY", "TEXTBEFORE", "TEXTAFTER", "TEXTSPLIT",
    "VSTACK", "HSTACK", "TOCOL", "TOROW", "WRAPROWS", "WRAPCOLS", "TAKE", "DROP", "CHOOSEROWS", "CHOOSECOLS",
    "EXPAND", "GROUPBY", "PIVOTBY", "PERCENTOF", "TRIMRANGE", "REGEXTEST", "REGEXEXTRACT", "REGEXREPLACE",
    "ARRAYTOTEXT", "VALUETOTEXT", "IMAGE",
}
_XLFN = re.compile(r"_xlfn\.(?:_xlws\.)?([A-Za-z][A-Za-z0-9.]*)\s*\(", re.I)
PREEXISTING = " (also in the original file)"


def _engine_gaps(wb, errors: dict) -> dict:
    """#NAME? cells that may come from functions the engine lacks, with the explanation.

    A cell qualifies when its own formula uses one of NEWER_FUNCTIONS and refers to
    no missing sheet, or when it only inherits #NAME? from such a cell.
    """
    gaps: dict = {}
    pending = {k: _formula(_value(wb[k[0]], k[1], k[2])) for k, v in errors.items()
               if v == "#NAME?" and k[0] in wb.sheetnames}
    for key, f in pending.items():
        newer = sorted({n.upper() for n in _XLFN.findall(f or "")} & NEWER_FUNCTIONS)
        if newer and all(rng[0] in wb.sheetnames for _, rng in _refs(f, key[0])):
            gaps[key] = (f"the formula uses {', '.join(newer)}, which some LibreOffice versions do not support, so the "
                         "error may come from the checking engine; check this cell in Excel or a newer LibreOffice "
                         "before changing it")

    def reads(f, sheet, key):
        return any(s == key[0] and c1 <= key[2] <= c2 and r1 <= key[1] <= (r2 if r2 is not None else key[1])
                   for _, (s, c1, r1, c2, r2) in _refs(f, sheet))

    changed = True
    while changed:
        changed = False
        for key, f in pending.items():
            if key in gaps or not f:
                continue
            source = next((g for g in gaps if reads(f, key[0], g)), None)
            if source is not None:
                gaps[key] = (f"it reads {_ref(*source)}, whose #NAME? may come from the checking engine "
                             "(see that cell)")
                changed = True
    return gaps


def check_errors(wb, values, findings, reason: str | None = None, before_values=None):
    if values is None:
        findings.append(Finding("warning", "recalculation-unavailable", "-",
                                f"{reason or 'LibreOffice is not available'}; formulas were not recalculated, so error "
                                "values (#REF!, #NAME?, #VALUE!, ...) were not checked. Make a working LibreOffice available "
                                "(soffice on PATH, or FAILBOOK_SOFFICE set to it), or open the file in a spreadsheet app "
                                "and look for them."))
        return
    errors = {}
    for ws in values.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.strip().upper() in ERRORS:
                    errors[(ws.title, c.row, c.column)] = c.value.strip().upper()
    gaps = _engine_gaps(wb, errors)
    for key, err in errors.items():
        sheet, row, col = key
        f = _formula(_value(wb[sheet], row, col)) if sheet in wb.sheetnames else None
        message = f"evaluates to {err}" + (f" ({f})" if f else "")
        old = _value(before_values[sheet], row, col) if before_values is not None and sheet in before_values.sheetnames else None
        if key in gaps:
            findings.append(Finding("warning", "error-value", _ref(sheet, row, col), f"{message}; {gaps[key]}"))
        elif isinstance(old, str) and old.strip().upper() == err:
            findings.append(Finding("warning", "error-value", _ref(sheet, row, col), message + PREEXISTING))
        else:
            findings.append(Finding("error", "error-value", _ref(sheet, row, col), message))


def _summary_rows(ws, col: int, r1: int, r2: int, rows: list[int]) -> set[int]:
    """The rows among ``rows`` (below the range r1..r2) whose cell in ``col`` is a formula
    computed only from that range or from earlier such rows, like =AVERAGE(B2:B10) or =B11/9
    in a stacked Total/Average/Max block."""
    summary: set[int] = set()

    def inside(lo: int, hi: int) -> bool:
        if r1 <= lo and hi <= r2:
            return True
        return hi - lo < 1000 and all(r1 <= x <= r2 or x in summary for x in range(lo, hi + 1))

    for r in rows:
        g = _formula(_value(ws, r, col))
        refs = [rng for _, rng in _refs(g, ws.title)] if g else []
        if refs and all(s == ws.title and rr2 is not None and inside(rr1, rr2) for s, _c1, rr1, _c2, rr2 in refs):
            summary.add(r)
    return summary


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
                    # Total below a column range that stops short of the data above it. Rows that
                    # summarise the range themselves (a Total/Average/Max block) are not data.
                    if c1 == c.column and r2 < c.row - 1:
                        filled = [r for r in range(r2 + 1, c.row) if _has_value(ws, r, c1)]
                        summary = _summary_rows(ws, c1, r1, r2, filled)
                        gap = [r for r in filled if r not in summary]
                        if gap:
                            rows_text = f"rows {gap[0]}-{gap[-1]}" if gap[-1] - gap[0] == len(gap) - 1 else \
                                "rows " + ", ".join(map(str, gap[:8])) + (", ..." if len(gap) > 8 else "")
                            findings.append(Finding("error", "range-omits-adjacent-cells", where,
                                                    f"{text} stops at row {r2}, but {rows_text} between it and the formula hold values"))
                    # A sum over a range that contains subtotals of the same range.
                    if agg in ("SUM", "PRODUCT"):
                        inner = []
                        for r in range(r1, r2 + 1):
                            g = _formula(_value(ws, r, c1))
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
def _blank(v) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


def _mostly_same(b_ws, a_ws, cells) -> bool:
    """True when at least half of the cells (row, col) that are filled before or after hold the same value."""
    pairs = [(_value(b_ws, r, c), _value(a_ws, r, c)) for r, c in cells]
    pairs = [(x, y) for x, y in pairs if x is not None or y is not None]
    return not pairs or 2 * sum(1 for x, y in pairs if x == y) >= len(pairs)


def _renames(before, after) -> tuple[dict[str, str], dict[str, tuple[str, float]]]:
    """Sheets that were renamed (old -> new), matched by content: at least half of the cells equal,
    so a renamed sheet that was also edited still matches. Also returns, for every other removed
    sheet, the most similar new sheet and its share of equal cells."""
    def content(ws):
        return {(c.row, c.column): c.value for row in ws.iter_rows() for c in row if c.value is not None}

    def similarity(old, new):
        if not old or not new:
            return 0.0
        return sum(1 for k, v in old.items() if new.get(k) == v) / max(len(old), len(new))

    removed = [n for n in before.sheetnames if n not in after.sheetnames]
    added = [n for n in after.sheetnames if n not in before.sheetnames]
    old_c = {n: content(before[n]) for n in removed}
    new_c = {n: content(after[n]) for n in added}
    scored = sorted(((similarity(old_c[r], new_c[a]), i, j, r, a) for i, r in enumerate(removed) for j, a in enumerate(added)),
                    key=lambda t: (-t[0], t[1], t[2]))
    renames: dict[str, str] = {}
    for score, _i, _j, r, a in scored:
        if score >= 0.5 and r not in renames and a not in renames.values():
            renames[r] = a
    closest = {}
    for score, _i, _j, r, a in scored:
        if r not in renames and r not in closest and score > 0:
            closest[r] = (a, score)
    return renames, closest


def compare(before, after, findings):
    renames, closest = _renames(before, after)
    for name in before.sheetnames:
        if name in after.sheetnames:
            continue
        renamed_to = renames.get(name)
        if renamed_to is None:
            hint = ""
            if name in closest:
                other, score = closest[name]
                hint = f" (the most similar new sheet, '{other}', has only {score:.0%} of its cells)"
            findings.append(Finding("error", "sheet-removed", name, f"sheet '{name}' existed before and is gone{hint}"))
            continue
        # A rename: fine, as long as nothing still points at the old name.
        pattern = re.compile(r"(?<![A-Za-z0-9_.])(?:'" + re.escape(name.replace("'", "''")) + "'|" + re.escape(name) + r")!")
        stale = [_ref(ws.title, c.row, c.column) for ws in after.worksheets for row in ws.iter_rows() for c in row
                 if _formula(c.value) and pattern.search(_formula(c.value))]
        stale += [f"name {dn}" for dn, d in after.defined_names.items() if pattern.search(str(d.attr_text or ""))]
        if stale:
            findings.append(Finding("error", "stale-sheet-reference", stale[0],
                                    f"sheet '{name}' was renamed to '{renamed_to}', but {len(stale)} formula(s)/name(s) still refer to '{name}'"))
    # Sheets present before and after (under the same or a new name).
    pairs = [(n, n) for n in before.sheetnames if n in after.sheetnames] + list(renames.items())
    back = {a: b for b, a in renames.items()}
    for bname, name in pairs:
        b, a = before[bname], after[name]
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
    # Unchanged formulas whose referenced columns or rows now mean something else: an error when
    # the label is gone or moved to another column/row (cells shifted under the formula), a
    # warning when it was only renamed or the column/row was rewritten in place.
    def moved_to(b_ws, a_ws, axis, idx, old, new):
        """Where ``old`` now is, or where ``new`` was before, other than ``idx``; None if neither."""
        if axis == "column":
            a_at = {_value(a_ws, 1, i): i for i in range(min(a_ws.max_column, 500), 0, -1) if i != idx}
            b_at = {_value(b_ws, 1, i): i for i in range(min(b_ws.max_column, 500), 0, -1) if i != idx}
            name = get_column_letter
        else:
            a_at = {_value(a_ws, i, 1): i for i in range(min(a_ws.max_row, 10000), 0, -1) if i != idx}
            b_at = {_value(b_ws, i, 1): i for i in range(min(b_ws.max_row, 10000), 0, -1) if i != idx}
            name = str
        if not _blank(old) and old in a_at:
            return f"{old!r} is now in {axis} {name(a_at[old])}"
        if not _blank(new) and new in b_at:
            return f"{new!r} was in {axis} {name(b_at[new])} before"
        return None

    relabelled: dict[tuple, list[str]] = {}
    for bname, name in pairs:
        for row in after[name].iter_rows():
            for c in row:
                f = _formula(c.value)
                if not f or _formula(_value(before[bname], c.row, c.column)) != f:
                    continue
                where = _ref(name, c.row, c.column)
                for text, (sheet, c1, r1, c2, r2) in _refs(f, name):
                    bsheet = back.get(sheet, sheet)
                    if sheet not in after.sheetnames or bsheet not in before.sheetnames:
                        continue
                    b_ws, a_ws = before[bsheet], after[sheet]
                    checks = []
                    if c1 == c2 and (sheet != name or r2 is None or r2 - r1 >= 1) and r1 > 1:
                        last = min(r2 if r2 is not None else max(b_ws.max_row, a_ws.max_row), r1 + 5000)
                        checks.append(("column", c1, _value(b_ws, 1, c1), _value(a_ws, 1, c1),
                                       [(r, c1) for r in range(r1, last + 1)], "header"))
                    if r2 == r1 and not (sheet == name and r1 == c.row):
                        width = min(max(b_ws.max_column, a_ws.max_column), 500)
                        checks.append(("row", r1, _value(b_ws, r1, 1), _value(a_ws, r1, 1),
                                       [(r1, col) for col in range(2, width + 1)], "label in column A"))
                    broken = False
                    for axis, idx, old, new, cells, part in checks:
                        if old == new:
                            continue
                        same = _mostly_same(b_ws, a_ws, cells)
                        shift = None if same else moved_to(b_ws, a_ws, axis, idx, old, new)
                        if _blank(new) or shift:
                            what = f"column {get_column_letter(idx)}" if axis == "column" else f"row {idx}"
                            findings.append(Finding("error", "reference-meaning-changed", where,
                                                    f"{f} refers to {what} of '{sheet}', whose {part} was {old!r} and is now "
                                                    f"{new!r}" + (f" ({shift})" if shift else "")))
                            broken = True
                            break
                        relabelled.setdefault((axis, sheet, idx, old, new, same), []).append(where)
                    if broken:
                        break
    for (axis, sheet, idx, old, new, same), wheres in relabelled.items():
        wheres = list(dict.fromkeys(wheres))
        what = f"column {get_column_letter(idx)}" if axis == "column" else f"row {idx}"
        part = "header" if axis == "column" else "label in column A"
        lead = f"{len(wheres)} unchanged formula(s) (e.g. at {wheres[0]}) refer to {what} of '{sheet}', whose {part} changed from {old!r} to {new!r}"
        if same:
            findings.append(Finding("warning", "reference-label-changed", wheres[0],
                                    f"{lead} while the other cells stayed the same; fine if it was renamed on purpose"))
        else:
            findings.append(Finding("warning", "reference-meaning-changed", wheres[0],
                                    f"{lead}, and most of the cells they read changed too; check that they still read what they should"))


# -- entry points ---------------------------------------------------------------------------
def _static_checks(wb, path: Path) -> list[Finding]:
    found: list[Finding] = []
    check_formulas(wb, found)
    check_ranges(wb, path, found)
    check_text_numbers(wb, found)
    check_pasted_totals(wb, found)
    return found


def audit(after: Path, before: Path | None = None, recalc: bool = True,
          recalculated_copy: Path | None = None) -> list[Finding]:
    """Findings for the edited workbook ``after``; ``before`` (the original) adds the comparison.

    ``recalculated_copy``, if given, receives the LibreOffice-recalculated copy of ``after``.
    """
    findings: list[Finding] = []
    after = Path(after)
    wb = openpyxl.load_workbook(after)
    old = openpyxl.load_workbook(before) if before is not None else None
    if recalc:
        try:
            values, reason = recalculated(after, keep=recalculated_copy), None
        except RecalcUnavailable as exc:
            values, reason = None, str(exc)
        old_values = openpyxl.load_workbook(before, data_only=True) if before is not None else None  # cached results
        check_errors(wb, values, findings, reason=reason, before_values=old_values)
    else:
        findings.append(Finding("warning", "recalculation-skipped", "-",
                                "recalculation was skipped, so error values (#REF!, #NAME?, ...) were not checked"))
    static = _static_checks(wb, after)
    if old is not None:
        # A symptom the original already shows was not caused by the edit.
        known = {(f.check, f.where, f.message) for f in _static_checks(old, Path(before))}
        for f in static:
            if (f.check, f.where, f.message) in known:
                f.level = "warning"
                f.message += PREEXISTING
    findings += static
    if old is not None:
        compare(old, wb, findings)
    return findings


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    p.add_argument("after", help="the edited workbook")
    p.add_argument("--before", help="the original workbook (enables collateral-damage checks)")
    p.add_argument("--recalculated", metavar="COPY.xlsx",
                   help="also save the copy LibreOffice recalculated (read it with openpyxl data_only=True)")
    p.add_argument("--no-recalc", action="store_true", help="skip LibreOffice recalculation")
    p.add_argument("--json", action="store_true")
    args = p.parse_args(argv)
    if args.recalculated and args.no_recalc:
        p.error("--recalculated needs recalculation; drop --no-recalc")
    findings = audit(Path(args.after), Path(args.before) if args.before else None, recalc=not args.no_recalc,
                     recalculated_copy=Path(args.recalculated) if args.recalculated else None)
    if args.json:
        print(json.dumps([asdict(f) for f in findings], ensure_ascii=False, indent=2))
    else:
        if not findings:
            print("no findings")
        for f in findings:
            print(f"[{f.level}] {f.check} at {f.where}: {f.message}")
        if args.recalculated and any(f.check == "recalculation-unavailable" for f in findings):
            print(f"no recalculated copy was written to {args.recalculated}")
        elif args.recalculated:
            print(f"recalculated copy: {args.recalculated}")
    return 1 if any(f.level == "error" for f in findings) else 0


if __name__ == "__main__":
    sys.exit(main())
