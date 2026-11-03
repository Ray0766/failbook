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
means check it. A symptom that the original file (--before) already shows is
reported as a warning marked "also in the original file": the edit did not
cause it, but the task may still need it changed.

Exit codes: 0 no error-level finding (warnings may remain), 1 at least one
error, 2 invalid command line, 3 a workbook could not be read (missing,
truncated, password-protected or not an .xlsx/.xlsm file; nothing was checked).

Recalculation forces LibreOffice to recompute every formula (a plain
`soffice --convert-to xlsx` keeps the results already cached in the file).
--recalculated saves that recalculated copy, to read with
openpyxl.load_workbook(COPY, data_only=True). With --no-recalc, or when
LibreOffice is unavailable, error values are not checked at all (the values
cached in the file are not read either).

This file is self-contained (Python 3.10+, openpyxl; LibreOffice optional for
recalculation) so that it can be shipped inside an Agent Skill. It is the
canonical copy; fixes/spreadsheet-verify-skill/skill/scripts/audit_xlsx.py is
generated from it by `failbook docs`.
"""

from __future__ import annotations

import argparse
import bisect
import datetime as dt
import json
import os
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import zipfile
from collections import Counter
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
    if row < 1 or col < 1:
        return None
    cells = getattr(ws, "_cells", None)
    if cells is None:
        return ws.cell(row, col).value
    c = cells.get((row, col))
    return None if c is None else c.value


def _filled(ws, min_row: int = 1) -> list:
    """The cells of ``ws`` that hold a value, in reading order.

    ``ws.iter_rows()`` visits, and creates, every position of the used range, which one formatted
    but empty cell far down the sheet stretches to a million rows; this only visits stored cells."""
    cells = getattr(ws, "_cells", None)
    if cells is None:
        return [c for row in ws.iter_rows(min_row=min_row) for c in row if c.value is not None]
    return [c for (r, _col), c in sorted(cells.items()) if r >= min_row and c.value is not None]


def _extent(ws) -> tuple[int, int]:
    """(last row, last column) holding a value; (0, 0) for an empty sheet."""
    rows = cols = 0
    for c in _filled(ws):
        rows, cols = max(rows, c.row), max(cols, c.column)
    return rows, cols


def _worksheets(wb) -> list[str]:
    """Names of the sheets that hold cells (chart sheets have none)."""
    return [ws.title for ws in wb.worksheets]


def _has_value(ws, row: int, col: int) -> bool:
    v = _value(ws, row, col)
    return v is not None and not (isinstance(v, str) and not v.strip())


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


class Unreadable(Exception):
    """A workbook could not be opened, so nothing was checked; the message says which file and why."""


_OLE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # compound-file signature: encrypted .xlsx or legacy .xls


def _load(path: Path, role: str, **kwargs):
    try:
        return openpyxl.load_workbook(path, **kwargs)
    except Exception as exc:
        try:
            with open(path, "rb") as fh:
                head = fh.read(8)
        except OSError as os_exc:
            raise Unreadable(f"cannot open {role} {path}: {os_exc.strerror or os_exc}") from None
        if head.startswith(_OLE):
            why = ("it is an OLE compound file, not a zip-based .xlsx: a password-protected (encrypted) workbook or "
                   "a legacy .xls file; save an unencrypted .xlsx copy and audit that")
        elif isinstance(exc, zipfile.BadZipFile) and head.startswith(b"PK"):
            why = "it is not a complete zip archive; the file may be truncated or damaged (for example by an interrupted save)"
        elif isinstance(exc, zipfile.BadZipFile):
            why = "it is not a zip archive, so it is not an .xlsx/.xlsm file"
        else:
            why = f"openpyxl could not read it ({type(exc).__name__}: {exc})"
        raise Unreadable(f"cannot open {role} {path}: {why}") from None


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
               if v == "#NAME?" and k[0] in _worksheets(wb)}
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
        for c in _filled(ws):
            if isinstance(c.value, str) and c.value.strip().upper() in ERRORS:
                errors[(ws.title, c.row, c.column)] = c.value.strip().upper()
    gaps = _engine_gaps(wb, errors)
    for key, err in errors.items():
        sheet, row, col = key
        f = _formula(_value(wb[sheet], row, col)) if sheet in _worksheets(wb) else None
        message = f"evaluates to {err}" + (f" ({f})" if f else "")
        old = (_value(before_values[sheet], row, col)
               if before_values is not None and sheet in _worksheets(before_values) else None)
        if key in gaps:
            findings.append(Finding("warning", "error-value", _ref(sheet, row, col), f"{message}; {gaps[key]}"))
        elif isinstance(old, str) and old.strip().upper() == err:
            findings.append(Finding("warning", "error-value", _ref(sheet, row, col), message + PREEXISTING))
        else:
            findings.append(Finding("error", "error-value", _ref(sheet, row, col), message))


def _column_aggregate(ws, row: int, col: int):
    """(FUNCTION, first row, last row, formula) when the cell is an aggregate such as =SUM(B2:B13)
    of one range in its own column that ends above it (a total row); None otherwise."""
    f = _formula(_value(ws, row, col))
    if not f:
        return None
    head = f[1:].lstrip("+").upper()
    agg = next((a for a in AGGREGATES if head.startswith(a + "(")), None)
    refs = [rng for _, rng in _refs(f, ws.title)]
    if agg is None or len(refs) != 1:
        return None
    sheet, c1, r1, c2, r2 = refs[0]
    if sheet != ws.title or c1 != col or c2 != col or r2 is None or r2 >= row:
        return None
    return agg, r1, r2, f


def _adds_above(ws, row: int, col: int) -> bool:
    """True when the cell is a formula that only adds cells above it in its own column, like
    =B2+B3+B4+B5 or =SUM(B2:B4)+SUM(B5:B7): a total written without one aggregate."""
    f = _formula(_value(ws, row, col))
    if not f:
        return False
    cells = 0
    for tok in _tokens(f):
        if tok.type == Token.OPERAND and tok.subtype == Token.RANGE:
            rng = parse_range(tok.value, ws.title)
            if not rng or rng[0] != ws.title or rng[1] != col or rng[3] != col or rng[4] is None or rng[4] >= row:
                return False
            cells += rng[4] - rng[2] + 1
        elif tok.type == Token.OP_IN and tok.value == "+":
            continue
        elif tok.type == Token.FUNC and tok.value.upper() in ("SUM(", ")"):
            continue
        elif tok.type in (Token.PAREN, Token.WSPACE) or (tok.type == Token.SEP and tok.subtype == Token.ARG):
            continue
        else:
            return False
    return cells >= 2


def _summary_rows(ws, col: int, r1: int, r2: int, rows: list[int]) -> set[int]:
    """The rows among ``rows`` (below the range r1..r2) whose cell in ``col`` is a formula
    computed only from that range or from earlier such rows, like =AVERAGE(B2:B10) or =B11/9
    in a stacked Total/Average/Max block, or an aggregate of the column above it (a Total
    row, or H1/H2 sums below a Total)."""
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
        elif _column_aggregate(ws, r, col):
            summary.add(r)
    return summary


def _closing_total(ws, col: int, r1: int, gap: list[int], below: int):
    """A total row between the last omitted row and ``below`` whose aggregate covers rows r1 to the
    last omitted row: the data block ends there, and the formula below it sums part of it."""
    for s in range(gap[-1] + 1, below):
        agg = _column_aggregate(ws, s, col)
        if agg and agg[1] <= r1 and agg[2] >= gap[-1]:
            return s, agg
    return None


def _splits_block(ws, col: int, row: int, agg: str) -> bool:
    """True when the ``agg`` formula in ``row`` is one of consecutive parts that split the data
    block above them exactly, like H1 =SUM(B2:B7) and H2 =SUM(B8:B13) below data ending in row 13.

    The parts sit in a run of adjacent aggregate rows of the column (a covering total such as
    Year =SUM(B2:B13) may be among them, above or below the parts, or absent), and the block they
    split ends right above that run (blank spacer rows aside)."""
    first = row
    while first > 1 and row - first < 100 and _column_aggregate(ws, first - 1, col):
        first -= 1
    end = first - 1
    while end > 1 and first - end <= 3 and not _has_value(ws, end, col):
        end -= 1
    if not _has_value(ws, end, col) or _column_aggregate(ws, end, col):
        return False
    run = []
    for t in range(first, first + 100):
        a = _column_aggregate(ws, t, col)
        if not a:
            break
        if a[0] == agg and a[2] <= end:
            run.append((a[1], a[2], t))
    # A part that contains another part is a covering total, not a part.
    parts = sorted(p for p in run if not any(q is not p and p[0] <= q[0] and q[1] <= p[1] for q in run))
    if len(parts) < 2 or parts[-1][1] != end or row not in {t for _lo, _hi, t in parts}:
        return False
    return all(b[0] == a[1] + 1 for a, b in zip(parts, parts[1:]))


def _grown_total(b_ws, positions: dict, col: int, row: int, f: str, r2: int, closing) -> tuple[int, str] | None:
    """Evidence from the original that a partial aggregate is stale: the same formula ``f`` (its
    range ending at ``r2``) was there before, below a total whose range also ended at ``r2``, and
    that total (``closing``, now ending lower) was extended while ``f`` was not. Returns (row, text)
    of the original total."""
    _t, (_agg, lo, hi, _text) = closing
    rows = [r for r, cc in positions.get(f, []) if cc == col]
    if hi <= r2 or not rows:
        return None
    s0 = min(rows, key=lambda r: (abs(r - row), r))
    for u in range(s0 - 1, max(r2, s0 - 100), -1):
        a = _column_aggregate(b_ws, u, col)
        if a and a[1] == lo and a[2] == r2:
            return u, a[3]
    return None


def _formula_positions(cells) -> dict[str, list[tuple[int, int]]]:
    """{formula text: [(row, column), ...]} of the formula cells among ``cells``."""
    positions: dict[str, list[tuple[int, int]]] = {}
    for c in cells:
        f = _formula(c.value)
        if f:
            positions.setdefault(f, []).append((c.row, c.column))
    return positions


def _sum_index(ws) -> dict[int, tuple[list[int], dict[int, list[tuple[int, int]]]]]:
    """{column: (sorted rows, {row: [(first, last) row of each range]})} of the =SUM(...) formulas
    of ``ws``, with the ranges of that same column that each one adds."""
    index: dict = {}
    for c in _filled(ws):
        g = _formula(c.value)
        if not g or not g[1:].upper().startswith("SUM("):
            continue
        spans = [(rr1, rr2) for _, (s, cc1, rr1, _cc2, rr2) in _refs(g, ws.title)
                 if s == ws.title and cc1 == c.column and rr2 is not None]
        rows, at = index.setdefault(c.column, ([], {}))
        rows.append(c.row)
        at[c.row] = spans
    for rows, _at in index.values():
        rows.sort()
    return index


def check_formulas(wb, findings, before=None):
    """Formula symptoms in ``wb``; ``before`` ({sheet name: the original's worksheet}) lets a stale
    partial aggregate below an extended total be told from a deliberate one."""
    before = before or {}
    for ws in wb.worksheets:
        b_ws, positions, sums = before.get(ws.title), None, None
        for c in _filled(ws):
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
                    # H1/H2 or quarters that split the block above them exactly are parts, not omissions.
                    if gap and not _splits_block(ws, c1, c.row, agg):
                        rows_text = f"rows {gap[0]}-{gap[-1]}" if gap[-1] - gap[0] == len(gap) - 1 else \
                            "rows " + ", ".join(map(str, gap[:8])) + (", ..." if len(gap) > 8 else "")
                        message = f"{text} stops at row {r2}, but {rows_text} between it and the formula hold values"
                        # Below a total that covers the omitted rows, a shorter range may be deliberate,
                        # unless the original shows that it used to end with that total.
                        closing = _closing_total(ws, c1, r1, gap, c.row)
                        if closing is None:
                            findings.append(Finding("error", "range-omits-adjacent-cells", where, message))
                        else:
                            if b_ws is not None and positions is None:
                                positions = _formula_positions(_filled(b_ws))
                            grown = _grown_total(b_ws, positions, c1, c.row, f, r2, closing) if b_ws is not None else None
                            if grown:
                                findings.append(Finding("error", "range-omits-adjacent-cells", where,
                                                        f"{message}; the total in row {closing[0]} was extended to "
                                                        f"them ({closing[1][3]}), but this formula was not, although "
                                                        f"in the original file both ended at row {r2} (row {grown[0]}: "
                                                        f"{grown[1]})"))
                            else:
                                findings.append(Finding("warning", "range-omits-adjacent-cells", where,
                                                        f"{message}; row {closing[0]} ({closing[1][3]}) covers them, "
                                                        "so check that the shorter range is deliberate"))
                # A sum over a range that contains subtotals of the same range.
                if agg in ("SUM", "PRODUCT"):
                    sums = _sum_index(ws) if sums is None else sums
                    rows, spans = sums.get(c1, ([], {}))
                    inner = [r for r in rows[bisect.bisect_left(rows, r1):bisect.bisect_right(rows, r2)]
                             if any(r1 <= rr1 and rr2 <= r2 for rr1, rr2 in spans[r])]
                    if inner:
                        findings.append(Finding("error", "double-counting", where,
                                                f"{text} includes rows {', '.join(map(str, inner[:6]))} which are themselves sums of rows in the same range"))


def check_ranges(wb, path, findings):
    """Charts, data validation and conditional formatting that stop before the data ends.

    The data ends at a blank cell or at a total row, which a chart or a highlighting rule normally
    leaves out: a row where any covered column holds an aggregate of the column above it, such as
    =SUM(B2:B7), or a formula that only adds cells above it, such as =B2+B3+B4."""
    def total_row(ws, row, c1, c2):
        return any(_column_aggregate(ws, row, col) or _adds_above(ws, row, col)
                   for col in range(c1, min(c2, c1 + 50) + 1))

    def short(ws, c1, r1, c2, r2, what, where):
        totals: dict[int, bool] = {}
        for col in range(c1, c2 + 1):
            if not _has_value(ws, r2, col):
                continue
            end = r2
            while _has_value(ws, end + 1, col):
                if end + 1 not in totals:
                    totals[end + 1] = total_row(ws, end + 1, c1, c2)
                if totals[end + 1]:
                    break
                end += 1
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
    names = _worksheets(wb)
    for chart in _charts(path):
        for s in chart["series"]:
            rng = parse_range(s.get("val") or "", chart["sheet"])
            if rng and rng[0] in names and rng[4]:
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
                    title_el = root.find("c:chart/c:title", ns)  # the chart's own title, not its axis titles
                    title = ""
                    if title_el is not None:
                        title = "".join(t.text or "" for t in title_el.iter(f"{{{ns['a']}}}t"))
                        title = title or title_el.findtext("c:tx/c:strRef/c:f", "", ns)
                    series = [{"val": (ser.find("c:val/c:numRef/c:f", ns).text if ser.find("c:val/c:numRef/c:f", ns) is not None else None)}
                              for ser in root.iter(f"{{{ns['c']}}}ser")]
                    found.append({"sheet": sh.get("name"), "title": title, "series": series})
    return found


def check_text_numbers(wb, findings):
    for ws in wb.worksheets:
        cols: dict[int, list] = {}
        for c in _filled(ws, min_row=2):
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
        for c in _filled(ws):
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


def _relative(text: str, axis: str) -> bool:
    """True when the reference's row (axis "row") or column (axis "column") has no $ anchor."""
    part = text.rsplit("!", 1)[-1]
    return not re.search(r"\$\d" if axis == "row" else r"\$[A-Za-z]", part)


def _spread(lo: int, hi: int, k: int) -> list[int]:
    """``lo..hi`` when it has at most ``k`` numbers, else ``k`` of them evenly spaced, ends included."""
    n = hi - lo + 1
    if n <= k:
        return list(range(lo, hi + 1))
    return sorted({lo + (n - 1) * i // (k - 1) for i in range(k)})


def _range_evidence(b_ws, a_ws, rng, dr: int, dc: int, last_row: int, sample: int = 200) -> tuple[int, int]:
    """(moved, stayed): referenced cells whose old value is now at the shifted / the original position.

    A cell whose current value is the same at both positions (a column of 1s, a repeated VAT rate)
    tells the two apart in no way and is not counted. Whole columns end at the original's last filled
    row, and large ranges are read at ``sample`` evenly spaced rows (and at most 50 columns)."""
    _sheet, c1, r1, c2, r2 = rng
    moved = stayed = 0
    for r in _spread(r1, min(r2, last_row) if r2 is not None else last_row, sample):
        for col in _spread(c1, c2, 50):
            old = _value(b_ws, r, col)
            if _blank(old):
                continue
            here, there = _value(a_ws, r, col), _value(a_ws, r + dr, col + dc)
            if here == there:
                continue
            if here == old:
                stayed += 1
            elif there == old:
                moved += 1
    return moved, stayed


def _moved_with(b_ws, a_ws, refs, dr: int, dc: int, last_row: int | None = None, cache: dict | None = None) -> bool:
    """True when the cells a moved formula reads moved by the same offset (dr, dc): their old values
    are now at the shifted position and other values sit at the referenced one, so the unchanged
    references are stale. When those cells stayed in place (a row or column was inserted between
    them and the formula), unchanged references are what a spreadsheet application keeps.

    ``cache`` (one per pair of sheets) keeps the evidence of each range, so formulas sharing a range
    (=SUMIF(A:A,A2,B:B) filled down) read it once."""
    last_row = _extent(b_ws)[0] if last_row is None else last_row
    cache = {} if cache is None else cache
    moved = stayed = 0
    for rng in refs:
        key = (rng, dr, dc)
        if key not in cache:
            cache[key] = _range_evidence(b_ws, a_ws, rng, dr, dc, last_row)
        m, s = cache[key]
        moved, stayed = moved + m, stayed + s
    return moved > 0 and moved >= stayed


def _renames(before, after) -> tuple[dict[str, str], dict[str, tuple[str, float]]]:
    """Sheets that were renamed (old -> new), matched by content: at least half of the old sheet's
    values (counted with repeats, wherever they are) are found in the new sheet, so a renamed sheet
    that was also extended, sorted or given a title row still matches. Also returns, for every
    other removed sheet, the most similar new sheet and that share. Constants are compared with
    constants; when either sheet holds none (a sheet of formulas), formula texts with formula texts."""
    def content(ws):
        values = [c.value for c in _filled(ws)]
        return (Counter(v for v in values if not _formula(v)), Counter(_formula(v) for v in values if _formula(v)))

    def similarity(old, new):
        (old_k, old_f), (new_k, new_f) = old, new
        old, new = (old_k, new_k) if old_k and new_k else (old_f, new_f)
        total = sum(old.values())
        return sum((old & new).values()) / total if total and new else 0.0

    old_names, new_names = _worksheets(before), _worksheets(after)
    removed = [n for n in old_names if n not in new_names]
    added = [n for n in new_names if n not in old_names]
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


def _chart_identity(chartsheet, renames: dict[str, str]) -> tuple[set[str], frozenset]:
    """(titles, series columns) of the charts on a chart sheet: the title texts, and the
    (sheet, first column, last column) of every series' values, with renamed sheets under their new
    name. Row extents are left out, so a series extended to new rows is still the same chart."""
    titles, columns = set(), set()
    for chart in getattr(chartsheet, "_charts", []):
        try:
            tx = chart.title.tx if chart.title is not None else None
            if tx is not None and tx.rich is not None:
                titles.add("".join(r.t or "" for p in tx.rich.p for r in (p.r or [])).strip())
            elif tx is not None and tx.strRef is not None:
                titles.add(tx.strRef.f)
        except AttributeError:
            pass
        for s in getattr(chart, "series", []):
            ref = getattr(getattr(getattr(s, "val", None), "numRef", None), "f", None)
            rng = parse_range(ref or "", "")
            if rng:
                columns.add((renames.get(rng[0], rng[0]), rng[1], rng[3]))
    return titles - {""}, frozenset(columns)


def compare(before, after, findings):
    renames, closest = _renames(before, after)
    old_names, new_names = _worksheets(before), _worksheets(after)
    for name in old_names:
        if name in new_names:
            continue
        renamed_to = renames.get(name)
        if renamed_to is None:
            hint = ""
            if name in closest:
                other, score = closest[name]
                hint = f" (the most similar new sheet, '{other}', holds only {score:.0%} of its values)"
            findings.append(Finding("error", "sheet-removed", name, f"sheet '{name}' existed before and is gone{hint}"))
            continue
        # A rename: fine, as long as nothing still points at the old name.
        pattern = re.compile(r"(?<![A-Za-z0-9_.])(?:'" + re.escape(name.replace("'", "''")) + "'|" + re.escape(name) + r")!")
        stale = [_ref(ws.title, c.row, c.column) for ws in after.worksheets for c in _filled(ws)
                 if _formula(c.value) and pattern.search(_formula(c.value))]
        stale += [f"name {dn}" for dn, d in after.defined_names.items() if pattern.search(str(d.attr_text or ""))]
        if stale:
            findings.append(Finding("error", "stale-sheet-reference", stale[0],
                                    f"sheet '{name}' was renamed to '{renamed_to}', but {len(stale)} formula(s)/name(s) still refer to '{name}'"))
    # Chart sheets hold no cells. One that is gone counts as renamed only when a new chart sheet shows
    # the same chart (the same title, or series over the same columns); otherwise it was removed.
    old_charts, new_charts = [cs.title for cs in before.chartsheets], [cs.title for cs in after.chartsheets]
    added = {cs.title: _chart_identity(cs, {}) for cs in after.chartsheets if cs.title not in old_charts}
    for cs in before.chartsheets:
        if cs.title in new_charts:
            continue
        name, (titles, columns) = cs.title, _chart_identity(cs, renames)
        same = next((n for n, (t, cols) in added.items() if (titles & t) or (columns and columns == cols)), None)
        if same is not None:
            del added[same]
            continue
        hint = (f" (the new chart sheet{'s' if len(added) > 1 else ''} "
                f"{', '.join(repr(n) for n in added)} show{'' if len(added) > 1 else 's'} another chart: "
                "neither the title nor the columns its series read match)") if added else ""
        findings.append(Finding("error", "sheet-removed", name, f"chart sheet '{name}' existed before and is gone{hint}"))
    # Sheets present before and after (under the same or a new name).
    pairs = [(n, n) for n in old_names if n in new_names] + list(renames.items())
    back = {a: b for b, a in renames.items()}
    for bname, name in pairs:
        b, a = before[bname], after[name]
        b_cells, a_cells = _filled(b), _filled(a)
        fb = sum(1 for c in b_cells if _formula(c.value))
        fa = sum(1 for c in a_cells if _formula(c.value))
        if fb >= 3 and fa < fb / 2:
            findings.append(Finding("error", "formulas-lost", name, f"'{name}' had {fb} formulas and now has {fa}"))
        if len(b.data_validations.dataValidation) > len(a.data_validations.dataValidation):
            findings.append(Finding("error", "validation-lost", name, "data validation rules were removed"))
        if b.freeze_panes and not a.freeze_panes:
            findings.append(Finding("warning", "frozen-panes-lost", name, f"frozen panes at {b.freeze_panes} were removed"))
        heads_b = {str(c.value): c for c in b_cells if c.row == 1}
        for c in [x for x in a_cells if x.row == 1]:
            old = heads_b.get(str(c.value))
            if old is not None and bool(old.font and old.font.b) and not bool(c.font and c.font.b):
                findings.append(Finding("warning", "header-style-lost", _ref(name, 1, c.column), f"header '{c.value}' lost its bold style"))
                break
        # Formula text that moved to another row or column while the cells it reads moved with it:
        # its references were not updated.
        positions_b = _formula_positions(b_cells)
        moved = []
        last_b, evidence = max((x.row for x in b_cells), default=0), {}
        for c in a_cells:
            f = _formula(c.value)
            if not f or "$" in f or f not in positions_b or (c.row, c.column) in positions_b[f]:
                continue
            refs = [rng for _, rng in _refs(f, name) if rng[0] == name]
            if not refs:
                continue
            r0, c0 = min(positions_b[f], key=lambda p: (abs(p[0] - c.row) + abs(p[1] - c.column), p))
            if _moved_with(b, a, refs, c.row - r0, c.column - c0, last_b, evidence):
                moved.append((c, (r0, c0)))
        if moved:
            c, (r0, c0) = moved[0]
            findings.append(Finding("error", "formula-moved-as-text", _ref(name, c.row, c.column),
                                    f"{len(moved)} formula(s) were moved without updating their references, although the "
                                    f"cells they read moved with them, e.g. {c.value} was at {get_column_letter(c0)}{r0}"))
    # Unchanged formulas whose referenced columns or rows now mean something else: an error when
    # the label is gone or moved to another column/row (cells shifted under the formula), a
    # warning when it was only renamed or the column/row was rewritten in place.
    cache: dict = {}

    def extent(ws):
        key = ("extent", id(ws))
        if key not in cache:
            cache[key] = _extent(ws)
        return cache[key]

    def labels(ws, axis):
        """{label: [positions]} of the header row (axis "column") or of column A (axis "row")."""
        key = ("labels", id(ws), axis)
        if key not in cache:
            at: dict = {}
            for c in _filled(ws):
                if (c.row == 1) if axis == "column" else (c.column == 1):
                    at.setdefault(c.value, []).append(c.column if axis == "column" else c.row)
            cache[key] = at
        return cache[key]

    def label(ws, axis, i):
        return _value(ws, 1, i) if axis == "column" else _value(ws, i, 1)

    def shift_of(b_ws, a_ws, axis, idx):
        """How far the label of column/row ``idx`` moved (0: still there), or None if it is not found
        at a position that held another label before."""
        old = label(b_ws, axis, idx)
        if _blank(old):
            return None
        if label(a_ws, axis, idx) == old:
            return 0
        found = [j for j in labels(a_ws, axis).get(old, []) if j != idx and label(b_ws, axis, j) != old]
        return min(found, key=lambda j: (abs(j - idx), j)) - idx if found else None

    def row_values(ws):
        """{row: (its non-blank constants, in column order)}: what identifies a row's content."""
        key = ("rows", id(ws))
        if key not in cache:
            rows: dict = {}
            for c in _filled(ws):
                if not _formula(c.value) and not _blank(c.value):
                    rows.setdefault(c.row, []).append(c.value)
            index: dict = {}
            for r, values in rows.items():
                index.setdefault(tuple(values), []).append(r)
            cache[key] = ({r: tuple(v) for r, v in rows.items()}, index)
        return cache[key]

    def unique_row(ws, values):
        """The one row of ``ws`` holding exactly ``values`` (two or more constants), else None."""
        found = row_values(ws)[1].get(values, []) if values and len(values) >= 2 else []
        return found[0] if len(found) == 1 else None

    def origin(b_ws, a_ws, idx):
        """The row of the original whose content row ``idx`` now holds: found by the row's constants
        when they are unique in both files, else by a column-A label that is unique in both; None
        when the row is new or cannot be told apart."""
        values = row_values(a_ws)[0].get(idx)
        if values and len(values) >= 2:
            return unique_row(b_ws, values) if unique_row(a_ws, values) == idx else None
        new = label(a_ws, "row", idx)
        if _blank(new):
            return None
        was, now = labels(b_ws, "row").get(new, []), labels(a_ws, "row").get(new, [])
        return was[0] if len(was) == 1 and now == [idx] else None

    def same_neighbour(b_ws, a_ws, idx, own):
        """True when the referenced row and the formula's own row hold content that was the same
        distance apart in the original: a relative reference refilled after rows were inserted or
        deleted between other rows still reads the same neighbour."""
        o_idx, o_own = origin(b_ws, a_ws, idx), origin(b_ws, a_ws, own)
        return o_idx is not None and o_own is not None and idx - o_idx == own - o_own

    def content_moved(b_ws, a_ws, idx):
        """Where the content of row ``idx`` of the original now is, when it is now at another row
        (found by its constants, unique in both files) and row ``idx`` holds something else."""
        values = row_values(b_ws)[0].get(idx)
        if unique_row(b_ws, values) != idx or row_values(a_ws)[0].get(idx) == values:
            return None
        j = unique_row(a_ws, values)
        return j if j is not None and j != idx else None

    def preview(values):
        shown = ", ".join(repr(v) for v in values[:4])
        return f"({shown}{', ...' if len(values) > 4 else ''})"

    def moved_to(b_ws, a_ws, axis, idx, old, new):
        """Where ``old`` now is, or where ``new`` was before, other than ``idx``; None if neither."""
        name = get_column_letter if axis == "column" else str
        k = shift_of(b_ws, a_ws, axis, idx)
        if k:
            return f"{old!r} is now in {axis} {name(idx + k)}"
        if not _blank(new):
            for j in labels(b_ws, axis).get(new, []):
                if j != idx and label(a_ws, axis, j) != new:
                    return f"{new!r} was in {axis} {name(j)} before"
        return None

    def line_stats(b_ws, a_ws, sheet, axis, idx):
        """(changed, compared, span): how many of the other filled cells of the column/row differ."""
        key = ("stats", sheet, axis, idx)
        if key not in cache:
            if axis == "column":
                last = min(max(extent(b_ws)[0], extent(a_ws)[0]), 5001)
                cells = [(r, idx) for r in range(2, last + 1)]
                span = f"{get_column_letter(idx)}2:{get_column_letter(idx)}{last}" if last >= 2 else ""
            else:
                width = min(max(extent(b_ws)[1], extent(a_ws)[1]), 500)
                cells = [(idx, col) for col in range(2, width + 1)]
                span = f"B{idx}:{get_column_letter(width)}{idx}" if width >= 2 else ""
            pairs_ = [(_value(b_ws, r, c), _value(a_ws, r, c)) for r, c in cells]
            pairs_ = [(x, y) for x, y in pairs_ if x is not None or y is not None]
            cache[key] = (sum(1 for x, y in pairs_ if x != y), len(pairs_), span)
        return cache[key]

    relabelled: dict[tuple, list[str]] = {}
    for bname, name in pairs:
        for c in _filled(after[name]):
            f = _formula(c.value)
            if not f or _formula(_value(before[bname], c.row, c.column)) != f:
                continue
            where = _ref(name, c.row, c.column)
            for text, (sheet, c1, r1, c2, r2) in _refs(f, name):
                bsheet = back.get(sheet, sheet)
                if sheet not in new_names or bsheet not in old_names:
                    continue
                b_ws, a_ws = before[bsheet], after[sheet]
                checks = []
                if c1 == c2 and (sheet != name or r2 is None or r2 - r1 >= 1) and r1 > 1:
                    checks.append(("column", c1, c.column, _value(b_ws, 1, c1), _value(a_ws, 1, c1), "header"))
                if r2 == r1 and not (sheet == name and r1 == c.row):
                    checks.append(("row", r1, c.row, _value(b_ws, r1, 1), _value(a_ws, r1, 1), "label in column A"))
                broken = False
                for axis, idx, own, old, new, part in checks:
                    # A relative reference on the formula's own sheet whose own row/column moved by the
                    # same offset as the one it reads (a filled-down running balance or month-on-month
                    # change after inserted or deleted rows) still reads the same neighbour.
                    neighbour = sheet == name and _relative(text, axis)
                    if old == new:
                        # Same label (a repeated category, or no label): the row's content may still
                        # have moved, as in a sort. Checked for references to another sheet or anchored
                        # ones; a relative one on its own sheet is most likely a filled-down pattern.
                        j = content_moved(b_ws, a_ws, idx) if axis == "row" and not neighbour else None
                        if j is not None:
                            now = row_values(a_ws)[0].get(idx)
                            findings.append(Finding(
                                "error", "reference-meaning-changed", where,
                                f"{f} refers to row {idx} of '{sheet}', whose values "
                                f"{preview(row_values(b_ws)[0][idx])} are now in row {j}; row {idx} now holds "
                                + (preview(now) if now else "no constants")))
                            broken = True
                            break
                        continue
                    if neighbour:
                        if axis == "row" and same_neighbour(b_ws, a_ws, idx, own):
                            continue
                        k = shift_of(b_ws, a_ws, axis, idx)
                        if k and shift_of(b_ws, a_ws, axis, own) == k:
                            continue
                    shift = moved_to(b_ws, a_ws, axis, idx, old, new)
                    if _blank(new) or shift:
                        what = f"column {get_column_letter(idx)}" if axis == "column" else f"row {idx}"
                        findings.append(Finding("error", "reference-meaning-changed", where,
                                                f"{f} refers to {what} of '{sheet}', whose {part} was {old!r} and is now "
                                                f"{new!r}" + (f" ({shift})" if shift else "")))
                        broken = True
                        break
                    relabelled.setdefault((axis, sheet, idx, old, new, line_stats(b_ws, a_ws, sheet, axis, idx)), []).append(where)
                if broken:
                    break
    for (axis, sheet, idx, old, new, (changed, compared, span)), wheres in relabelled.items():
        wheres = list(dict.fromkeys(wheres))
        what = f"column {get_column_letter(idx)}" if axis == "column" else f"row {idx}"
        part = "header" if axis == "column" else "label in column A"
        lead = f"{len(wheres)} unchanged formula(s) (e.g. at {wheres[0]}) refer to {what} of '{sheet}', whose {part} changed from {old!r} to {new!r}"
        if not compared:
            findings.append(Finding("warning", "reference-label-changed", wheres[0],
                                    f"{lead}; the rest of that {axis} is empty; fine if it was renamed on purpose"))
        elif not changed:
            findings.append(Finding("warning", "reference-label-changed", wheres[0],
                                    f"{lead} while the other cells of {span} stayed the same; fine if it was renamed on purpose"))
        else:
            findings.append(Finding("warning", "reference-meaning-changed", wheres[0],
                                    f"{lead}, and {changed} of the {compared} other filled cells of {span} changed too; "
                                    "check that they still read what they should"))


# -- entry points ---------------------------------------------------------------------------
def _static_checks(wb, path: Path, before=None) -> list[Finding]:
    """The checks of the edited file alone; ``before`` (the original workbook) only sharpens them."""
    found: list[Finding] = []
    originals = {}
    if before is not None:
        renames = _renames(before, wb)[0]
        back = {new: old for old, new in renames.items()}
        originals = {n: before[back.get(n, n)] for n in _worksheets(wb) if back.get(n, n) in _worksheets(before)}
    check_formulas(wb, found, originals)
    check_ranges(wb, path, found)
    check_text_numbers(wb, found)
    check_pasted_totals(wb, found)
    return found


def audit(after: Path, before: Path | None = None, recalc: bool = True,
          recalculated_copy: Path | None = None) -> list[Finding]:
    """Findings for the edited workbook ``after``; ``before`` (the original) adds the comparison.

    ``recalculated_copy``, if given, receives the LibreOffice-recalculated copy of ``after``.
    Raises Unreadable when ``after`` or ``before`` cannot be opened.
    """
    findings: list[Finding] = []
    after = Path(after)
    wb = _load(after, "the edited workbook")
    old = _load(before, "the original workbook (--before)") if before is not None else None
    if recalc:
        try:
            values, reason = recalculated(after, keep=recalculated_copy), None
        except RecalcUnavailable as exc:
            values, reason = None, str(exc)
        # The original's cached results, to tell error values the edit caused from ones it already had.
        old_values = _load(before, "the original workbook (--before)", data_only=True) if before is not None else None
        check_errors(wb, values, findings, reason=reason, before_values=old_values)
    else:
        findings.append(Finding("warning", "recalculation-skipped", "-",
                                "recalculation was skipped, so error values (#REF!, #NAME?, ...) were not checked"))
    static = _static_checks(wb, after, before=old)
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


#: Help texts shared with ``failbook audit``.
NO_RECALC_HELP = ("skip LibreOffice recalculation; error values (#REF!, #DIV/0!, ...) are then not checked, "
                  "not even the values cached in the file")
EXIT_CODES = ("exit codes: 0 no error-level finding (warnings may remain), 1 at least one error, "
              "2 invalid command line, 3 a workbook could not be read (missing, truncated, password-protected "
              "or not an .xlsx/.xlsm file; nothing was checked)")


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description=__doc__.split("\n\n")[0], epilog=EXIT_CODES)
    p.add_argument("after", help="the edited workbook")
    p.add_argument("--before", help="the original workbook (enables collateral-damage checks)")
    p.add_argument("--recalculated", metavar="COPY.xlsx",
                   help="also save the copy LibreOffice recalculated (read it with openpyxl data_only=True)")
    p.add_argument("--no-recalc", action="store_true", help=NO_RECALC_HELP)
    p.add_argument("--json", action="store_true")
    args = p.parse_args(argv)
    if args.recalculated and args.no_recalc:
        p.error("--recalculated needs recalculation; drop --no-recalc")
    try:
        findings = audit(Path(args.after), Path(args.before) if args.before else None, recalc=not args.no_recalc,
                         recalculated_copy=Path(args.recalculated) if args.recalculated else None)
    except Unreadable as exc:
        print(f"audit not run: {exc}", file=sys.stderr)
        return 3
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
