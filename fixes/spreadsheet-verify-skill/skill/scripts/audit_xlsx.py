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
import functools
import json
import os
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import unicodedata
import zipfile
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path

import openpyxl
from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import column_index_from_string, get_column_letter

ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}
AGGREGATES = {"SUM", "AVERAGE", "COUNT", "COUNTA", "MIN", "MAX", "SUBTOTAL", "PRODUCT"}
SUMMARY_WORDS = re.compile(r"summary|total|dashboard|report|汇总|合计", re.I)
NUMBER_TEXT = re.compile(r"^\s*[-+(]?\s*[$€£¥]?\s*\d{1,3}(,\d{3})+(\.\d+)?\s*\)?\s*$|^\s*[-+]?[$€£¥]?\d+\.\d+\s*%?\s*$|^\s*[-+]?\d+\s*%\s*$")
DATE_TEXT = re.compile(r"^\s*\d{4}-\d{2}-\d{2}\s*$")
#: What a formula's label says about the position of the row it reads ('Top row', 'Highest price', '#2',
#: '最高'). After a sort such a formula may be meant to read whatever row now holds that position, so
#: compare() turns a moved reference into a warning when the edited table shows the row holds it
#: (_position_claims gives the claims, compare's positional() checks them).
_NUMBERS = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9,
            "ten": 10, "first": 1, "second": 2, "third": 3, "fourth": 4, "fifth": 5, "sixth": 6, "seventh": 7,
            "eighth": 8, "ninth": 9, "tenth": 10, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5, "六": 6,
            "七": 7, "八": 8, "九": 9, "十": 10}
_NUM = r"(\d+|one|two|three|four|five|six|seven|eight|nine|ten|[一二两三四五六七八九十]+)"
_ORDINAL = r"first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth"


def _words(alternatives: str) -> str:
    """English words on ASCII edges: Chinese text next to an English word is no word boundary."""
    return r"(?<![a-z])(?:" + alternatives + r")(?![a-z])"


TOP_LIST = re.compile(_words(r"top|first|leading|best") + r"[\s-]*" + _NUM + r"(?![a-z\d])|前\s*" + _NUM)
BOTTOM_LIST = re.compile(_words(r"bottom|last|worst") + r"[\s-]*" + _NUM + r"(?![a-z\d])|后\s*" + _NUM + r"\s*名")
#: A rank: (pattern, counted from the bottom).
RANKS = [(re.compile(r"倒数第\s*" + _NUM), True),
         (re.compile(r"第\s*" + _NUM), False),
         (re.compile(r"#\s*(\d+)(?!\d)"), False),
         (re.compile(_words(r"no\.?|number|rank(?:ed)?|place|position|pos\.?") + r"\s*#?\s*" + _NUM + r"(?![a-z\d])"), False),
         (re.compile(r"(?<![a-z\d])(\d+)(?:st|nd|rd|th)(?![a-z])"), False),
         (re.compile(_words("(" + _ORDINAL + ")")), False)]
MEDALS = re.compile(_words(r"runner[- ]?up|gold|silver|bronze") + r"|冠军|亚军|季军|[金银铜]牌|🥇|🥈|🥉")
MEDAL_RANK = {"gold": 1, "silver": 2, "bronze": 3, "冠军": 1, "亚军": 2, "季军": 3, "金牌": 1, "银牌": 2, "铜牌": 3,
              "🥇": 1, "🥈": 2, "🥉": 3}
TOP_WORDS = re.compile(_words(r"top|best|leader|leading|winner|winning|champion|bestsell\w*") + r"|榜首|首位|首行|头名")
BOTTOM_WORDS = re.compile(_words(r"bottom|worst") + r"|末位|垫底|倒数")
EITHER_WORDS = re.compile(_words(r"fastest|quickest|slowest") + r"|最快|最慢")
HIGH_WORDS = re.compile(_words(r"highest|largest|biggest|greatest|longest|max|maximum|most|dearest|priciest|peak")
                        + r"|最高|最大|最多|最贵|最长")
LOW_WORDS = re.compile(_words(r"lowest|smallest|shortest|min|minimum|least|cheapest|fewest") + r"|最低|最小|最少|最便宜|最短")
#: Words that turn 'biggest' into the lowest number ('Biggest drop').
INVERT_WORDS = re.compile(_words(r"drops?|falls?|fallers?|declines?|decreases?|loss(?:es)?|losers?") + r"|跌|降")
NEW_WORDS = re.compile(_words(r"latest|newest|most\s+recent|recent|last") + r"|最新|最近|最后")
OLD_WORDS = re.compile(_words(r"earliest|oldest|least\s+recent") + r"|最早|最老|最先")
#: Fixed phrases and units that hold a position word but name none ('First name', 'Q1' as 'first quarter',
#: '(min)' for minutes, 'last 12 months').
NOT_POSITION = re.compile(
    _words(r"(?:first|last|middle|given|family)\s*names?|(?:first|second|third|fourth|last|1st|2nd|3rd|4th)\s+"
           r"(?:quarter|half|qtr)|best\s+before|first\s+aid|first[- ]time(?:rs?)?|top[- ]?ups?"
           r"|last\s+(?:year|month|week|quarter|updated|modified|login|seen)"
           r"|(?:last|past|next|first)\s+\d+\s+(?:days?|weeks?|months?|years?|quarters?)")
    + r"|\(\s*(?:min|mins|minutes)\s*\)|(?<=\d)\s*mins?(?![a-z])|per\s+min(?![a-z])"
    + r"|第[一二两三四五六七八九十\d]+(?:季度|车间|分店|分公司|期|周|月|年|季|批|版)")
#: A name with a number ('Warehouse #1', 'PO #3', 'Acct. No. 2'), unless the word is a position word.
NUMBERED_NAME = re.compile(r"(?<![^\W\d_])([^\W\d_]{2,})\.?\s*(?:#|no\.?)\s*\d+")
RANK_HEADER = re.compile(r"^\s*(?:rank(?:ing)?|#|no\.?|pos(?:ition)?|place|standing|名次|排名|排行|位次)\s*$", re.I)
TOTAL_LABEL = re.compile(r"^\s*(?:grand\s+)?(?:total|sum|subtotal|average|avg|mean)(?![a-z])|^\s*(?:合计|总计|小计|总和|平均)",
                         re.I)
DATE_LIKE = re.compile(r"^\s*(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})")
YEAR_HEADER = re.compile(r"(?<![a-z])(?:year|yr|fy)(?![a-z])|年", re.I)


def _number(text: str) -> int | None:
    """3 for '3', 'three', 'third' or '三'; 12 for '十二'; None when it is not a number."""
    t = text.strip().lower()
    if t.isascii() and t.isdigit():
        return int(t) if len(t) <= 6 else None
    if t in _NUMBERS:
        return _NUMBERS[t]
    if len(t) == 2 and t[0] == "十" and t[1] in _NUMBERS:
        return 10 + _NUMBERS[t[1]]
    if len(t) in (2, 3) and t[0] in _NUMBERS and t[1] == "十":
        return _NUMBERS[t[0]] * 10 + (_NUMBERS.get(t[2], 0) if len(t) == 3 else 0)
    return None


def _fold(text) -> str:
    """Text compared case-insensitively, with full-width digits and letters as ASCII."""
    return unicodedata.normalize("NFKC", str(text)).casefold()


def _is_position_word(word: str) -> bool:
    w = _fold(word)
    return w in {"rank", "ranked", "place", "position", "pos", "pick", "seed", "spot", "number", "no", "top", "bottom"} \
        or any(p.fullmatch(w) for p in (TOP_WORDS, BOTTOM_WORDS, HIGH_WORDS, LOW_WORDS, EITHER_WORDS, MEDALS)) \
        or re.fullmatch(_ORDINAL, w) is not None


@functools.lru_cache(maxsize=4096)
def _position_claims(text: str, field: str = "") -> tuple:
    """What the label ``text`` says about the position of the row its formula reads, as claims to check
    against the table: ('pos', k, from_bottom) for the k-th row, ('within', n, from_bottom) for one of the
    first or last n, ('value', 'max' | 'min', k) for the k-th highest or lowest number, ('date', 'max' |
    'min') for the newest or oldest date, ('edge', 'top' | 'bottom') for the first or last row or the
    highest or lowest number, and ('either',) for any of these. ``field`` is the header of the column the
    formula reads: in a label it is a field name ('Last service', 'Max load'), not a position."""
    t = _fold(text)
    f = re.sub(r"\(.*?\)", " ", _fold(field)).strip(" :-–—")
    if len(f) >= 2 and _position_claims(f) and f in t:
        t = t.replace(f, " ")
    t = NOT_POSITION.sub(" ", t)
    t = NUMBERED_NAME.sub(lambda m: m.group(0) if _is_position_word(m.group(1)) else " ", t)
    claims: list[tuple] = []
    used: list[tuple[int, int]] = []

    def free(m):
        return not any(a < m.end() and m.start() < b for a, b in used)

    def take(span, *found):
        used.append(span)
        claims.extend(found)

    for pattern, bottom in ((TOP_LIST, False), (BOTTOM_LIST, True)):
        for m in pattern.finditer(t):
            n = _number(next(g for g in m.groups() if g))
            if n and n < 1000:
                take(m.span(), ("within", n, bottom))
    for pattern, bottom in RANKS:
        for m in pattern.finditer(t):
            n = _number(next(g for g in m.groups() if g))
            if not (n and free(m) and n < 1000):
                continue
            after = t[m.end():m.end() + 24]
            value = re.match(r"[\s-]*(?:(" + HIGH_WORDS.pattern + r")|(" + LOW_WORDS.pattern + r")|([高大多])|([低小少]))", after)
            if value:
                high = bool(value.group(1) or value.group(3))
                take((m.start(), m.end() + value.end()), ("value", "max" if high else "min", n))
                continue
            if re.match(r"[\s-]*(?:to|from)[\s-]+(?:the\s+)?(?:last|bottom|end)", after):
                bottom = True
            take(m.span(), ("pos", n, bottom), *((("date", "min"),) if n == 1 and m.group(0).strip() == "first" else ()))
    for m in MEDALS.finditer(t):
        if free(m):
            w = re.sub(r"\s|-", "", m.group(0))
            take(m.span(), ("pos", MEDAL_RANK.get(w, 2), False))
    invert = INVERT_WORDS.search(t) is not None
    for m in NEW_WORDS.finditer(t):
        if free(m):
            take(m.span(), ("date", "max"), *((("pos", 1, True),) if m.group(0) in ("last", "最后") else ()))
    for m in OLD_WORDS.finditer(t):
        if free(m):
            take(m.span(), ("date", "min"))
    for pattern, high in ((HIGH_WORDS, True), (LOW_WORDS, False)):
        for m in pattern.finditer(t):
            if free(m):
                take(m.span(), ("value", "max" if high != invert else "min", 1))
    for pattern, claim in ((TOP_WORDS, ("edge", "top")), (BOTTOM_WORDS, ("edge", "bottom")), (EITHER_WORDS, ("either",))):
        for m in pattern.finditer(t):
            if free(m):
                take(m.span(), claim)
    return tuple(claims)


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
    =B2+B3+B4+B5 or =SUM(B2:B4)+SUM(B5:B7): a total written without one aggregate. The cells must
    be the whole block above it, from its first value (below a label, a blank or the header row,
    whatever the header holds) to the last one (blank spacer rows aside), so a column of =B5+B6
    steps (a computed series) is data, not totals."""
    f = _formula(_value(ws, row, col))
    if not f:
        return False
    spans = []
    for tok in _tokens(f):
        if tok.type == Token.OPERAND and tok.subtype == Token.RANGE:
            rng = parse_range(tok.value, ws.title)
            if not rng or rng[0] != ws.title or rng[1] != col or rng[3] != col or rng[4] is None or rng[4] >= row:
                return False
            spans.append((rng[2], rng[4]))
        elif tok.type == Token.OP_IN and tok.value == "+":
            continue
        elif tok.type == Token.FUNC and tok.value.upper() in ("SUM(", ")"):
            continue
        elif tok.type in (Token.PAREN, Token.WSPACE) or (tok.type == Token.SEP and tok.subtype == Token.ARG):
            continue
        else:
            return False
    spans.sort()
    if not spans:
        return False
    top, bottom = spans[0]
    for lo, hi in spans[1:]:
        if lo > bottom + 1:
            return False  # a hole between the added cells
        bottom = max(bottom, hi)
    if bottom - top < 1 or row - bottom > 4 or any(_has_value(ws, r, col) for r in range(bottom + 1, row)):
        return False
    above = _value(ws, top - 1, col)
    return top <= 2 or not (_formula(above) or (isinstance(above, (int, float)) and not isinstance(above, bool)))


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


def _overlapping_part(ws, col: int, row: int, agg: str, r1: int, r2: int, memo: dict):
    """(first row, last row, formula row) of another ``agg`` in the run of total rows holding ``row`` whose
    range overlaps r1-r2 without either containing the other (H1 =SUM(B2:B8), H2 =SUM(B8:B13))."""
    _end, run = _aggregate_run(ws, col, row, memo)
    for a, lo, hi, t in run:
        if (t != row and a == agg and lo <= r2 and hi >= r1
                and not (lo <= r1 and r2 <= hi) and not (r1 <= lo and hi <= r2)):
            return lo, hi, t
    return None


def _covering_in_run(ws, col: int, row: int, r1: int, gap: list[int], memo: dict):
    """An aggregate in the run of total rows holding ``row`` (above or below it) that covers rows r1 to
    the last omitted row, as (row, aggregate): a Year total below Q1..Q4 covers what a part leaves out."""
    _end, run = _aggregate_run(ws, col, row, memo)
    for a, lo, hi, t in run:
        if t != row and lo <= r1 and hi >= gap[-1]:
            return t, _aggregate_at(ws, t, col, memo)
    return None


def _closing_total(ws, col: int, r1: int, gap: list[int], below: int):
    """A total row between the last omitted row and ``below`` whose aggregate covers rows r1 to the
    last omitted row: the data block ends there, and the formula below it sums part of it."""
    for s in range(gap[-1] + 1, below):
        agg = _column_aggregate(ws, s, col)
        if agg and agg[1] <= r1 and agg[2] >= gap[-1]:
            return s, agg
    return None


def _aggregate_at(ws, row: int, col: int, memo: dict | None):
    """``_column_aggregate``, remembered in ``memo``."""
    if memo is None:
        return _column_aggregate(ws, row, col)
    key = ("agg", ws.title, row, col)
    if key not in memo:
        memo[key] = _column_aggregate(ws, row, col)
    return memo[key]


def _aggregate_run(ws, col: int, row: int, memo: dict | None = None) -> tuple[int | None, list[tuple[str, int, int, int]]]:
    """(end, run) for the run of adjacent aggregate rows of the column that holds ``row``: ``end`` is
    the last row of the data block right above the run (blank spacer rows aside), and ``run`` the
    (function, first row, last row, row) of each aggregate in it; (None, []) when no data block ends
    right above the run."""
    first = row
    while first > 1 and row - first < 100 and _aggregate_at(ws, first - 1, col, memo):
        first -= 1
    key = ("run", ws.title, col, first)
    if memo is not None and key in memo:
        return memo[key]
    end = first - 1
    while end > 1 and first - end <= 3 and not _has_value(ws, end, col):
        end -= 1
    if not _has_value(ws, end, col) or _aggregate_at(ws, end, col, memo):
        result = (None, [])
    else:
        run = []
        for t in range(first, first + 100):
            a = _aggregate_at(ws, t, col, memo)
            if not a:
                break
            run.append((a[0], a[1], a[2], t))
        result = (end, run)
    if memo is not None:
        memo[key] = result
    return result


def _splits_block(ws, col: int, row: int, agg: str, memo: dict | None = None) -> bool:
    """True when the ``agg`` formula in ``row`` is one of consecutive parts that split the data
    block above them exactly, like H1 =SUM(B2:B7) and H2 =SUM(B8:B13) below data ending in row 13.

    The parts sit in a run of adjacent aggregate rows of the column (a covering total such as
    Year =SUM(B2:B13) may be among them, above or below the parts, or absent), and the block they
    split, from its first data row to its last, ends right above that run (blank spacer rows aside)."""
    info = _split_parts(ws, col, row, agg, memo if memo is not None else {})
    if info is None:
        return False
    top, end, parts, gaps, _covering = info
    return (not [g for g in gaps if g >= parts[0][0]] and parts[-1][1] == end
            and row in {t for _lo, _hi, t in parts})


def _first_rows(ws, memo: dict) -> dict[int, int]:
    """{column: the first row that holds a value in it}."""
    key = ("first", ws.title)
    if key not in memo:
        first: dict[int, int] = {}
        for c in _filled(ws):
            first.setdefault(c.column, c.row)
        memo[key] = first
    return memo[key]


def _is_data(ws, row: int, col: int, memo: dict | None = None) -> bool:
    """True when the cell holds a number, or a formula that is not a total of the column above it: a row
    of a data block. Row 1 is the header row, and so is the first filled cell of the column, even when
    they hold numbers (years as column headers, below a title row)."""
    if row <= 1 or (memo is not None and row <= _first_rows(ws, memo).get(col, 0)):
        return False
    v = _value(ws, row, col)
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return True
    return (bool(_formula(v)) and not _column_aggregate(ws, row, col) and not _adds_above(ws, row, col)
            and not _from_totals(ws, row, col) and not _wrapped_total(ws, row, col))


def _wrapped_total(ws, row: int, col: int) -> bool:
    """True for an aggregate of the column above wrapped in rounding, like =ROUND(SUM(B2:B13),0)."""
    f = _formula(_value(ws, row, col)) or ""
    m = re.match(r"(ROUND|ROUNDUP|ROUNDDOWN|INT|TRUNC|ABS)\((\w+)\(", f[1:].lstrip("+").upper().replace(" ", ""))
    if not m or m.group(2) not in AGGREGATES:
        return False
    refs = [rng for _, rng in _refs(f, ws.title)]
    return bool(refs) and all(sheet == ws.title and c1 == col and c2 == col and r2 is not None and r2 < row
                              for sheet, c1, _r1, c2, r2 in refs)


def _from_totals(ws, row: int, col: int) -> bool:
    """True when the formula reads only totals of its own column above it, like a Variance =B11-B6
    between two blocks: a summary row, not data."""
    refs = [rng for _, rng in _refs(_formula(_value(ws, row, col)) or "", ws.title)]
    return bool(refs) and all(
        sheet == ws.title and c1 == col and c2 == col and r2 is not None and r2 - r1 < 50 and r2 < row
        and all(_column_aggregate(ws, r, col) for r in range(r1, r2 + 1))
        for sheet, c1, r1, c2, r2 in refs)


def _block_top(ws, col: int, row: int, memo: dict) -> int:
    """The first row of the run of data cells (``_is_data``) in ``col`` that goes up from ``row``."""
    tops = memo.setdefault((ws.title, col), {})
    path, r = [], row
    while r not in tops:
        path.append(r)
        if _is_data(ws, r - 1, col, memo):
            r -= 1
        else:
            tops[r] = r
    for p in path:
        tops[p] = tops[r]
    return tops[r]


def _split_parts(ws, col: int, row: int, agg: str, memo: dict):
    """How the ``agg`` parts of the run of aggregate rows holding ``row`` cover the data block right
    above the run, or None when there is no such block or fewer than two parts.

    Returns (top, end, parts, gaps, covering): the block's rows ``top``-``end``; the parts, as sorted
    (first row, last row, formula row), leaving out an aggregate that contains another one (a covering
    total); the rows of the block in no part; and whether an ``agg`` of the run covers the whole block.
    Parts that overlap give None: they do not try to split the block."""
    end, aggregates = _aggregate_run(ws, col, row, memo)
    if end is None or not aggregates:
        return None
    key = ("parts", ws.title, col, aggregates[0][3], agg)
    if key in memo:
        return memo[key]
    memo[key] = _parts_of(ws, col, agg, end, aggregates, memo)
    return memo[key]


def _parts_of(ws, col: int, agg: str, end: int, aggregates, memo: dict):
    run = [(lo, hi, t) for a, lo, hi, t in aggregates if a == agg and hi <= end]
    parts = sorted(p for p in run if not any(q is not p and p[0] <= q[0] and q[1] <= p[1] for q in run))
    if len(parts) < 2 or any(b[0] <= a[1] for a, b in zip(parts, parts[1:])):
        return None
    top = min(_block_top(ws, col, end, memo), parts[0][0])
    covered = set()
    for lo, hi, _t in parts:
        covered.update(range(max(lo, top), hi + 1))
    gaps = [r for r in range(top, end + 1) if r not in covered and _is_data(ws, r, col, memo)]
    covering = any(a == agg and lo <= top and hi >= end for a, lo, hi, _t in aggregates)
    return top, end, parts, gaps, covering


def _runs(rows: list[int]) -> list[tuple[int, int]]:
    """Consecutive runs of sorted rows, as (first, last)."""
    out: list[tuple[int, int]] = []
    for r in rows:
        if out and r == out[-1][1] + 1:
            out[-1] = (out[-1][0], r)
        else:
            out.append((r, r))
    return out


def _rows_text(rows: list[int]) -> str:
    runs = _runs(rows)
    if len(rows) == 1:
        return f"row {rows[0]}"
    if len(runs) == 1:
        return f"rows {runs[0][0]}-{runs[0][1]}"
    return "rows " + ", ".join(str(r) for r in rows[:8]) + (", ..." if len(rows) > 8 else "")


def _near_split(ws, col: int, row: int, agg: str, memo: dict):
    """(top, end, parts, gaps) when the ``agg`` parts around ``row`` are meant to split the block above
    them (H1/H2, four quarters) but leave rows out: every gap is shorter than the shortest part. Gaps
    after the last part below a total of the whole block are left alone: sub-period parts of an
    earlier period (H1 2024 and H2 2024 kept below a total extended to 2025) leave those rows out on
    purpose; so are rows above the first part, unless that part is the short one. None otherwise."""
    info = _split_parts(ws, col, row, agg, memo)
    if info is None:
        return None
    top, end, parts, gaps, covering = info
    if not gaps or row not in {t for _lo, _hi, t in parts}:
        return None
    lengths = [hi - lo + 1 for lo, hi, _t in parts]
    # Rows above the first part are a gap only when that part is the short one (Q1 =SUM(B3:B4) beside
    # three-month quarters); otherwise they are an opening row or an earlier period, as they are when a
    # total of the whole block stands in the run.
    lead = [g for g in gaps if g < parts[0][0]]
    rest = [g for g in gaps if g >= parts[0][0]]
    keep_lead = bool(lead) and not covering and lengths[0] + len(lead) <= max(lengths[1:])
    if any(b - a + 1 >= min(lengths) for a, b in _runs(rest)):
        return None
    if covering:  # rows after the last part belong to a later period (H1 2024, H2 2024 below a 2025 total)
        rest = [g for g in rest if g <= parts[-1][1]]
    gaps = (lead if keep_lead else []) + rest
    return (top, end, parts, gaps) if gaps else None


def _groups_parts(ws, col: int, row: int, agg: str, r1: int, r2: int, memo: dict) -> bool:
    """True when rows r1-r2 are exactly some consecutive parts of an exact split of the block in the run
    of aggregate rows holding ``row`` (H1 =SUM(B2:B7) beside quarters Q1..Q4 of B2:B13)."""
    info = _split_parts(ws, col, row, agg, memo)
    if info is None:
        return False
    parts = info[2]
    inside = [p for p in parts if r1 <= p[0] and p[1] <= r2]
    return len(inside) >= 2 and any(p[0] == r1 for p in parts) and any(p[1] == r2 for p in parts)


def _inner_part(ws, col: int, row: int, agg: str, memo: dict) -> bool:
    """True when the formula in ``row`` is a part, not the last, of parts that follow each other
    without a gap from the first to the last: its range stops above the next part, not above data."""
    info = _split_parts(ws, col, row, agg, memo)
    if info is None:
        return False
    parts, gaps = info[2], info[3]
    rows = [t for _lo, _hi, t in parts]
    return (row in rows and row != parts[-1][2]
            and not [g for g in gaps if parts[0][0] <= g <= parts[-1][1]])


def _split_gap(sheet: str, col: int, agg: str, top: int, end: int, parts, gaps) -> Finding:
    """The error for parts that split rows ``top``-``end`` except ``gaps``, at the part next to the first gap."""
    letter = get_column_letter(col)
    span = lambda p: f"{letter}{p[0]}:{letter}{p[1]} (row {p[2]})"  # noqa: E731
    first = gaps[0]
    if first < parts[0][0]:
        part, how = parts[0], f"{span(parts[0])} starts at row {parts[0][0]}, below the first data row (row {top})"
    elif first > parts[-1][1]:
        part, how = parts[-1], f"{span(parts[-1])} ends at row {parts[-1][1]}, above the last data row (row {end})"
    else:
        i = next(i for i, p in enumerate(parts) if p[0] > first)
        part, how = parts[i], (f"{span(parts[i])} starts at row {parts[i][0]}, but the part before it ends at row "
                               f"{parts[i - 1][1]}")
    verb = "is" if len(gaps) == 1 else "are"
    rows = sorted(t for _lo, _hi, t in parts)
    where = f"rows {rows[0]}-{rows[-1]}" if rows[-1] - rows[0] == len(rows) - 1 else "rows " + ", ".join(map(str, rows))
    return Finding("error", "range-omits-adjacent-cells", _ref(sheet, part[2], col),
                   f"{len(parts)} {agg} parts in {where} split rows {top}-{end} of column {letter}, but "
                   f"{_rows_text(gaps)} {verb} in none of them: {how}")


def _late_start(ws, b_ws, col: int, row: int, agg: str, text: str, r1: int, r2: int, memo: dict, part: bool = False):
    """(level, message) when the ``agg`` range r1-r2 in ``row`` starts below the first row of the data
    block it sums (=SUM(B3:B13) over data from row 2), else None. Not for a part below an aggregate of
    the same rows that covers the left-out ones (a 'Q4 2024' below the year's total). An error when
    the column's own parts include the left-out rows while this formula contains one of them (a
    'Total' below 2024 / 2025 parts that leaves out January), or when the original file's formula in
    this cell covered them and they hold the same values; otherwise a warning, since a total may leave
    out a first row on purpose (an opening balance). For a ``part`` of an exact split (2024 beside 2025)
    only the original's evidence counts."""
    if not _is_data(ws, r1, col, memo):
        return None
    top = _block_top(ws, col, r1, memo)
    # A second header row (years below the column titles) has no label in column A, unlike the rows
    # the range sums.
    while top < r1 and _blank(_value(ws, top, 1)) and not _blank(_value(ws, r1, 1)):
        top += 1
    if top >= r1:
        return None
    omitted = list(range(top, r1))

    def says(rows):
        holds = "holds a value" if len(rows) == 1 else "hold values"
        return f"{text} starts at row {r1}, but {_rows_text(rows)} right above it in the same column {holds}"

    # The original's own formula here, or its aggregate with the same label in column A (a 'Total' that
    # moved down when rows were added), covered the rows, which hold the same values: an error.
    old = _column_aggregate(b_ws, row, col) if b_ws is not None else None
    if b_ws is not None and not old and not _blank(_value(ws, row, 1)):
        label = _value(ws, row, 1)
        old = next((a for r in range(1, min(_extent(b_ws)[0], row + 200) + 1)
                    if _value(b_ws, r, 1) == label and (a := _column_aggregate(b_ws, r, col))), None)
    if (old and old[0] == agg and old[1] <= top and old[2] >= r1
            and all(_value(b_ws, r, col) == _value(ws, r, col) for r in omitted)):
        return "error", f"{says(omitted)}; in the original file this formula covered them ({old[3]})"
    if part:
        return None
    _end, run = _aggregate_run(ws, col, row, memo)
    others = [(a, lo, hi, t) for a, lo, hi, t in run if t != row]
    if any(a == agg and lo <= top and hi >= r2 for a, lo, hi, _t in others):
        return None
    # The rows summed by the parts before it (H2 below H1) are not left out.
    before_it = set()
    for a, lo, hi, _t in others:
        if a == agg and hi < r1:
            before_it.update(range(lo, hi + 1))
    omitted = [r for r in omitted if r not in before_it]
    if not omitted:
        return None
    message = says(omitted)
    contains = any(a == agg and r1 <= lo and hi <= r2 for a, lo, hi, _t in others)
    inside = [(lo, hi, t) for a, lo, hi, t in others if a == agg and lo <= r1 - 1 and hi >= r1]
    if contains and inside:
        lo, hi, t = inside[0]
        letter = get_column_letter(col)
        return "error", (f"{message}, which the part {letter}{lo}:{letter}{hi} in row {t} includes, so this total "
                         "leaves out data its parts cover")
    return "warning", f"{message}; check that {'it' if len(omitted) == 1 else 'they'} should be left out"


def _grown_total(b_ws, ws, positions: dict, col: int, row: int, f: str, r1: int, r2: int, closing) -> tuple | None:
    """Evidence from the original that an aggregate below an extended total is stale: the same
    formula ``f`` (rows ``r1``-``r2``) was there before, below a total over exactly the same rows,
    and that total (``closing``, which starts at ``r1`` too and now ends lower) was extended while
    ``f`` was not. A part that starts lower (a 'Q4 2024' or 'H2 2024' kept below an all-time total)
    gives no evidence, as such parts also keep their range in correct edits.

    Returns (row, text) of the original total: the aggregate of those rows with the closing total's
    function, else the one whose column-A label is the closing total's label in ``ws`` (the edited
    sheet); (None, None) when the original has an aggregate of those rows but neither identifies the
    total (the closing total may be new); None without evidence."""
    t, (agg, lo, hi, _text) = closing
    rows = [r for r, cc in positions.get(f, []) if cc == col]
    if hi <= r2 or r1 != lo or not rows:
        return None
    s0 = min(rows, key=lambda r: (abs(r - row), r))
    found = []
    for u in range(s0 - 1, max(r2, s0 - 100), -1):
        a = _column_aggregate(b_ws, u, col)
        if a and a[1] == lo and a[2] == r2 and a[3] != f:
            found.append((u, a))
    if not found:
        return None
    name = _value(ws, t, 1)
    same = [(u, a) for u, a in found if a[0] == agg] or [(u, a) for u, a in found
                                                         if not _blank(name) and _value(b_ws, u, 1) == name]
    return (same[0][0], same[0][1][3]) if same else (None, None)


def _was_block_total(b_ws, positions: dict, col: int, row: int, f: str, r1: int, r2: int) -> int | None:
    """The row of the original where the same formula ``f`` was the total of the whole data block
    above it: its rows ``r1``-``r2`` were that block, from its first value (below a label or a blank)
    to its last one, which sat right above the formula (blank spacer rows aside)."""
    rows = [r for r, cc in positions.get(f, []) if cc == col and r > r2]
    if not rows:
        return None
    s0 = min(rows, key=lambda r: (abs(r - row), r))
    if any(_has_value(b_ws, r, col) for r in range(r2 + 1, s0)) or s0 - r2 > 4:
        return None
    above = _value(b_ws, r1 - 1, col)
    if r1 > 1 and (_formula(above) or (isinstance(above, (int, float)) and not isinstance(above, bool))):
        return None
    return s0 if all(_has_value(b_ws, r, col) for r in _spread(r1, r2, 200)) else None


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
    memo: dict = {}
    reported: set = set()
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
                own = c1 == c.column and r2 < c.row
                flagged = len(findings)
                # Parts meant to split the block above them (H1/H2, quarters) that leave rows out: one
                # error at the part next to the gap, naming the rows; the other parts are not omissions.
                near = _near_split(ws, c1, c.row, agg, memo) if own else None
                if near is not None:
                    top, end, parts, gaps = near
                    if (ws.title, c1, parts[0][2]) not in reported:
                        reported.add((ws.title, c1, parts[0][2]))
                        findings.append(_split_gap(ws.title, c1, agg, top, end, parts, gaps))
                # Total below a column range that stops short of the data above it. Rows that
                # summarise the range themselves (a Total/Average/Max block) are not data.
                elif c1 == c.column and r2 < c.row - 1:
                    filled = [r for r in range(r2 + 1, c.row) if _has_value(ws, r, c1)]
                    summary = _summary_rows(ws, c1, r1, r2, filled)
                    gap = [r for r in filled if r not in summary]
                    rows_text = _rows_text(gap) if gap else ""
                    holds = "holds a value" if len(gap) == 1 else "hold values"
                    message = f"{text} stops at row {r2}, but {rows_text} between it and the formula {holds}"
                    # H1/H2 or quarters that split the block above them exactly are parts, not omissions,
                    # unless the original shows that this formula was the total of the whole block,
                    # which has grown since (a stale total followed by a new sum of the added rows), and
                    # no aggregate in the same run totals the grown block (2024 / 2025 / Total).
                    if gap and _splits_block(ws, c1, c.row, agg, memo):
                        end, run = _aggregate_run(ws, c1, c.row, memo)
                        totalled = any(a == agg and lo <= r1 and hi >= end for a, lo, hi, _t in run)
                        if b_ws is not None and positions is None and not totalled:
                            positions = _formula_positions(_filled(b_ws))
                        was = (_was_block_total(b_ws, positions, c1, c.row, f, r1, r2)
                               if b_ws is not None and not totalled else None)
                        if was is not None:
                            findings.append(Finding("warning", "range-omits-adjacent-cells", where,
                                                    f"{message}; in the original file this formula (row {was}) was the "
                                                    f"total of the whole block, rows {r1}-{r2}, which has grown since, so "
                                                    "check that it should not cover the new rows"))
                    elif gap and (_groups_parts(ws, c1, c.row, agg, r1, r2, memo) or _inner_part(ws, c1, c.row, agg, memo)):
                        pass  # H1 =SUM(B2:B7) beside quarters: it adds Q1 and Q2; Q1 stops where Q2 starts
                    elif gap:
                        # Below a total that covers the omitted rows, a shorter range may be deliberate,
                        # unless the original shows that it used to end with that total.
                        closing = _closing_total(ws, c1, r1, gap, c.row) or _covering_in_run(ws, c1, c.row, r1, gap, memo)
                        twin = _overlapping_part(ws, c1, c.row, agg, r1, r2, memo)
                        if twin is not None:
                            letter = get_column_letter(c1)
                            both = list(range(max(r1, twin[0]), min(r2, twin[1]) + 1))
                            findings.append(Finding("error", "range-omits-adjacent-cells", where,
                                                    f"{message}, and the part {letter}{twin[0]}:{letter}{twin[1]} in row "
                                                    f"{twin[2]} overlaps it: both include {_rows_text(both)}, so the "
                                                    f"parts count {'it' if len(both) == 1 else 'them'} twice"))
                        elif closing is None:
                            findings.append(Finding("error", "range-omits-adjacent-cells", where, message))
                        else:
                            if b_ws is not None and positions is None:
                                positions = _formula_positions(_filled(b_ws))
                            grown = (_grown_total(b_ws, ws, positions, c1, c.row, f, r1, r2, closing)
                                     if b_ws is not None else None)
                            if grown and grown[0] is not None:
                                findings.append(Finding("error", "range-omits-adjacent-cells", where,
                                                        f"{message}; the total in row {closing[0]} was extended to "
                                                        f"them ({closing[1][3]}), but this formula was not, although "
                                                        f"in the original file both covered rows {r1}-{r2} (the total "
                                                        f"was in row {grown[0]}: {grown[1]})"))
                            elif grown:  # the original's total is not identified: the one here may be new
                                findings.append(Finding("error", "range-omits-adjacent-cells", where,
                                                        f"{message}; row {closing[0]} ({closing[1][3]}) covers them, "
                                                        "but this formula does not, although in the original file it "
                                                        f"covered the same rows ({r1}-{r2}) as another aggregate of "
                                                        "the column"))
                            else:
                                findings.append(Finding("warning", "range-omits-adjacent-cells", where,
                                                        f"{message}; row {closing[0]} ({closing[1][3]}) covers them, "
                                                        "so check that the shorter range is deliberate"))
                # A range that starts below the first row of the data block it sums.
                if own and near is None and len(findings) == flagged and _column_aggregate(ws, c.row, c.column):
                    start = _late_start(ws, b_ws, c1, c.row, agg, text, r1, r2, memo,
                                        part=_splits_block(ws, c1, c.row, agg, memo))
                    if start is not None:
                        findings.append(Finding(start[0], "range-omits-adjacent-cells", where, start[1]))
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
                    series = [{"val": next((el.text for el in (ser.find("c:val/c:numRef/c:f", ns), ser.find("c:yVal/c:numRef/c:f", ns))
                                            if el is not None), None)}  # yVal: scatter and bubble charts
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
    if r2 is None and dr and not dc:
        return 0, 0  # a whole column reads the same cells wherever its rows moved: no evidence either way
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
    them and the formula), unchanged references are what a spreadsheet application keeps. A whole
    column (B:B) reads the same cells after rows moved, so it is no evidence for a formula moved only
    up or down; a formula that reads only whole columns is never moved-as-text by a row shift alone.

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
    constants; only when the old sheet holds none (a sheet of formulas) are formula texts compared,
    so a sheet whose constants all vanished is removed even if its formula texts reappear elsewhere."""
    def content(ws):
        values = [c.value for c in _filled(ws)]
        return (Counter(v for v in values if not _formula(v)), Counter(_formula(v) for v in values if _formula(v)))

    def similarity(old, new):
        (old_k, old_f), (new_k, new_f) = old, new
        old, new = (old_k, new_k) if old_k else (old_f, new_f)
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


#: Marks a chart series over a defined name that is no plain range ("[" cannot start a sheet name).
NAMED = "[name]"


def _defined_name(wb, ref: str) -> tuple[str, str, str | None]:
    """(scope, name, text) of the defined name a chart series reads (``Data!Revenue``,
    ``[0]!Revenue`` or ``Revenue``). A name of the sheet itself comes first (scope: that sheet),
    then the workbook's (scope ""); ``text`` is what the name stands for, None when there is no such
    name (scope: the sheet as written)."""
    sheet, _, name = ref.rpartition("!")
    sheet = sheet[1:-1].replace("''", "'") if sheet.startswith("'") and sheet.endswith("'") else sheet
    scopes = [(sheet, getattr(wb[sheet], "defined_names", {}))] if sheet in _worksheets(wb) else []
    scopes.append(("", wb.defined_names))
    for scope, names in scopes:
        items = names.items() if hasattr(names, "items") else ((d.name, d) for d in getattr(names, "definedName", []))
        for key, d in items:
            if str(key).lower() == name.lower():
                return scope, name.lower(), d.attr_text
    return sheet, name.lower(), None


def _chart_identity(chartsheet, renames: dict[str, str], wb=None) -> tuple[set[str], frozenset]:
    """(titles, series columns) of the charts on a chart sheet: the title texts, and the
    (sheet, first column, last column) of every series' values (``val``, or ``yVal`` in scatter and
    bubble charts; ``xVal`` only for a series without either), with renamed sheets under their new
    name. Row extents are left out, so a series extended to new rows is still the same chart. A
    series over a defined name of ``wb`` counts by the range the name stands for or, when that is no
    plain range (an OFFSET formula), as (NAMED, scope, name), which pairs the same name only."""
    titles, columns = set(), set()

    def ref_of(data):
        for kind in ("numRef", "strRef"):
            f = getattr(getattr(data, kind, None), "f", None)
            if f:
                return f
        return None

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
            refs = [ref_of(getattr(s, kind, None)) for kind in ("val", "yVal")]
            refs = [r for r in refs if r] or [ref_of(getattr(s, "xVal", None))]
            for ref in refs:
                rng = parse_range(ref or "", "")
                if rng is None and ref and wb is not None and re.fullmatch(r"(?:.+!)?[A-Za-z_\\][\w.]*", ref):
                    scope, key, text = _defined_name(wb, ref)
                    rng = parse_range(text or "", "")
                    if rng is None:
                        columns.add((NAMED, renames.get(scope, scope), key))
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
    # Chart sheets hold no cells. One that is gone counts as renamed when a new chart sheet shows the
    # same chart: the same title, or series over at least one of the same columns or defined names (a
    # series may have been added or removed). It was removed when it has a title or a series over a
    # cell range and no new chart sheet matches. An old chart with neither (its series read no range,
    # or a name that is no plain range) may be any new chart sheet left over, which is a warning;
    # such charts are paired last, so they cannot take another chart's renamed sheet.
    old_charts, new_charts = [cs.title for cs in before.chartsheets], [cs.title for cs in after.chartsheets]
    added = {cs.title: _chart_identity(cs, {}, after) for cs in after.chartsheets if cs.title not in old_charts}
    gone = [(cs.title, _chart_identity(cs, renames, before)) for cs in before.chartsheets if cs.title not in new_charts]
    removed, unknown = [], []

    def known(identity):
        titles, columns = identity
        return bool(titles) or any(c[0] != NAMED for c in columns)

    for name, (titles, columns) in sorted(gone, key=lambda g: not known(g[1])):
        same = next((n for n, (t, cols) in added.items() if (titles & t) or (columns & cols)), None)
        if same is not None:
            del added[same]
        elif known((titles, columns)):
            removed.append(name)
        else:
            unknown.append(name)
    for name in removed:
        if len(added) > 1:
            hint = (f" (the new chart sheets {', '.join(repr(n) for n in added)} show other charts: none of them "
                    "shares a title, or a column or name that a series plots as values, with it)")
        elif added:
            hint = (f" (the new chart sheet {next(iter(added))!r} shows another chart: the two share no title and "
                    "no column or name that a series plots as values)")
        else:
            hint = ""
        findings.append(Finding("error", "sheet-removed", name, f"chart sheet '{name}' existed before and is gone{hint}"))
    for name in unknown:
        if not added:
            findings.append(Finding("error", "sheet-removed", name, f"chart sheet '{name}' existed before and is gone"))
            continue
        other = next(iter(added))
        del added[other]
        findings.append(Finding("warning", "sheet-removed", name,
                                f"chart sheet '{name}' is gone; the new chart sheet '{other}' may be it under a new "
                                "name, but the old chart has no title and its series read no cell range, so the audit "
                                "cannot tell; check that it was not deleted"))
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

    def row_values(ws, cols=None):
        """{row: (its non-blank constants, in column order)}: what identifies a row's content; with
        ``cols`` (first, last), only the constants in those columns (one table of the sheet)."""
        key = ("rows", id(ws), cols)
        if key not in cache:
            rows: dict = {}
            for c in _filled(ws):
                if cols is not None and not cols[0] <= c.column <= cols[1]:
                    continue
                if not _formula(c.value) and not _blank(c.value):
                    rows.setdefault(c.row, []).append(c.value)
            index: dict = {}
            for r, values in rows.items():
                index.setdefault(tuple(values), []).append(r)
            cache[key] = ({r: tuple(v) for r, v in rows.items()}, index)
        return cache[key]

    def unique_row(ws, values, cols=None):
        """The one row of ``ws`` holding exactly ``values`` (two or more constants), else None."""
        found = row_values(ws, cols)[1].get(values, []) if values and len(values) >= 2 else []
        return found[0] if len(found) == 1 else None

    def blocks(ws):
        """{column: (first, last)}: the run of columns holding a value that each column belongs to;
        runs are separated by columns that hold nothing (a table, and a summary beside it)."""
        key = ("blocks", id(ws))
        if key not in cache:
            out, run = {}, []
            for col in sorted({c.column for c in _filled(ws)}) + [None]:
                if run and (col is None or col != run[-1] + 1):
                    out.update(dict.fromkeys(run, (run[0], run[-1])))
                    run = []
                if col is not None:
                    run.append(col)
            cache[key] = out
        return cache[key]

    def beside(b_ws, a_ws, row, col, c1, c2):
        """True when the formula at (row, col) reads cells c1-c2 of its own row in a table that starts in
        column A, from another block of columns (an empty column between them in both files), and its
        own block holds a constant in that row that did not change, such as 'Apple price' next to
        =$C$2: a summary beside the table, not a formula of the table's row, which moves with it."""
        for ws in (b_ws, a_ws):
            bl = blocks(ws)
            if col not in bl or c1 not in bl or bl.get(c2) != bl[c1] or bl[col] == bl[c1] or bl[c1][0] != 1:
                return False
        lo, hi = blocks(a_ws)[col]
        for k in range(lo, min(hi, lo + 49) + 1):
            v = _value(a_ws, row, k)
            if (k != col and isinstance(v, str) and not _blank(v) and not _formula(v) and _value(b_ws, row, k) == v
                    and column_count(b_ws, k, v) == 1 and column_count(a_ws, k, v) == 1):
                return True
        return False

    def column_count(ws, k, v):
        key = ("colcounts", id(ws))
        if key not in cache:
            counts: dict = {}
            for c in _filled(ws):
                counts.setdefault(c.column, Counter())[c.value] += 1
            cache[key] = counts
        return cache[key].get(k, Counter())[v]

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
        if o_idx is None or o_own is None:
            return False
        if idx - o_idx == own - o_own:
            return True
        # Now next to each other, in the same order as before, and every row that was between them
        # is gone from the sheet: the rows between were deleted.
        lo, hi = sorted((o_idx, o_own))
        return (abs(idx - own) == 1 and (idx < own) == (o_idx < o_own) and hi - lo <= 1000
                and all(gone(b_ws, a_ws, r) for r in range(lo + 1, hi)))

    def gone(b_ws, a_ws, r):
        """True when row ``r`` of the original was blank, or held two or more constants (unique in
        the original) that no row of the edited sheet holds together any more."""
        values = row_values(b_ws)[0].get(r)
        if values is None:
            return not any(_has_value(b_ws, r, col) for col in range(1, min(extent(b_ws)[1], 500) + 1))
        return len(values) >= 2 and unique_row(b_ws, values) == r and not row_values(a_ws)[1].get(values)

    def new_row(b_ws, a_ws, idx):
        """True when row ``idx`` holds two or more constants, unique in the edited sheet, that no row
        of the original held together: a row that was inserted (or rewritten)."""
        values = row_values(a_ws)[0].get(idx)
        return (values is not None and len(values) >= 2 and unique_row(a_ws, values) == idx
                and not row_values(b_ws)[1].get(values))

    def content_moved(b_ws, a_ws, idx, cols=None):
        """Where the content of row ``idx`` of the original now is, when it is now at another row
        (found by its constants, unique in both files) and row ``idx`` holds something else. With
        ``cols``, only the constants of those columns count (the table the reference reads, not a
        summary beside it in the same rows)."""
        values = row_values(b_ws, cols)[0].get(idx)
        if unique_row(b_ws, values, cols) != idx or row_values(a_ws, cols)[0].get(idx) == values:
            return None
        j = unique_row(a_ws, values, cols)
        return j if j is not None and j != idx else None

    def filled_down(ws, row, col, f):
        """True when the cell above or below holds the same formula filled down its column (=D6+C7
        above =D7+C8): a running balance or change column, not a one-off formula."""
        key = ("filled", id(ws), row, col)
        if key not in cache:
            here, found = f"{get_column_letter(col)}{row}", False
            for r in (row - 1, row + 1):
                other = _formula(_value(ws, r, col))
                if other:
                    try:
                        found = Translator(f, origin=here).translate_formula(f"{get_column_letter(col)}{r}") == other
                    except Exception:  # a formula the translator cannot read is not evidence
                        found = False
                if found:
                    break
            cache[key] = found
        return cache[key]

    def emptied(b_ws, a_ws, row, c1, c2):
        """The text of the first cell of ``row``, columns c1..c2, that held something in the original,
        when every one of those cells is now empty; None otherwise."""
        cols = range(c1, min(c2, c1 + 49) + 1)
        if any(not _blank(_value(a_ws, row, col)) for col in cols):
            return None
        return next((f"{get_column_letter(col)}{row}, which it reads, held {_value(b_ws, row, col)!r}"
                     for col in cols if not _blank(_value(b_ws, row, col))), None)

    def preview(values):
        shown = ", ".join(repr(v) for v in values[:4])
        return f"({shown}{', ...' if len(values) > 4 else ''})"

    def moves(b_ws, a_ws, idx, old, new):
        """(went_to, came_from): the row the content of row ``idx`` moved to, and the row its new content
        came from, by labels in column A when each is unique in both files, else by the row's constants
        in the table's columns; None where it cannot be told."""
        was_at, now_at = labels(b_ws, "row"), labels(a_ws, "row")
        if was_at.get(old) == [idx] and len(now_at.get(old, [])) == 1 and now_at.get(new) == [idx] \
                and len(was_at.get(new, [])) == 1:
            return now_at[old][0], was_at[new][0]
        cols = blocks(b_ws).get(1)
        if cols is None:
            return None, None
        went = content_moved(b_ws, a_ws, idx, cols)
        now = row_values(a_ws, cols)[0].get(idx)
        return went, unique_row(b_ws, now, cols) if now else None

    def merged(ws):
        """({(row, col): (top row, left col)} for the cells of small merged ranges, [the large ranges])."""
        key = ("merged", id(ws))
        if key not in cache:
            cells, large = {}, []
            for m in ws.merged_cells.ranges:
                if (m.max_row - m.min_row + 1) * (m.max_col - m.min_col + 1) <= 5000:
                    for r in range(m.min_row, m.max_row + 1):
                        for k in range(m.min_col, m.max_col + 1):
                            cells[(r, k)] = (m.min_row, m.min_col)
                else:
                    large.append((m.min_row, m.max_row, m.min_col, m.max_col))
            cache[key] = cells, large
        return cache[key]

    def anchor(ws, row, col):
        """The cell that holds what (row, col) shows: the top-left cell of its merged range, or itself."""
        cells, large = merged(ws)
        if (row, col) in cells:
            return cells[(row, col)]
        for r1, r2, c1_, c2_ in large:
            if r1 <= row <= r2 and c1_ <= col <= c2_:
                return r1, c1_
        return row, col

    def cell_text(ws, row, col):
        """The constant shown at (row, col), the top-left value of a merged range included; None for a
        formula or a blank."""
        v = _value(ws, row, col)
        if v is None:
            v = _value(ws, *anchor(ws, row, col))
        return None if v is None or _formula(v) or _blank(v) else v

    def cell_name(ws, row, col):
        r, k = anchor(ws, row, col)
        return f"{get_column_letter(k)}{r}"

    def column_rows(ws, col):
        """(rows of column ``col`` showing text, rows holding any value), sorted."""
        key = ("column-rows", id(ws), col)
        if key not in cache:
            filled = {c.row for c in _filled(ws) if c.column == col and not _blank(c.value)}
            cells, large = merged(ws)
            shown = filled | {r for (r, k) in cells if k == col}
            for r1, r2, c1_, c2_ in large:
                if c1_ <= col <= c2_:
                    shown |= set(range(r1, min(r2, r1 + 100000) + 1))
            texts = sorted(r for r in shown if isinstance(cell_text(ws, r, col), str))
            cache[key] = texts, sorted(filled)
        return cache[key]

    def texts_above(ws, row, col, n=2):
        """[(cell, text)]: the nearest ``n`` texts within 39 rows above (row, col) in its column, over
        formulas and numbers, the second at most two blank rows above the first (a header, a title)."""
        texts, filled = column_rows(ws, col)
        i = bisect.bisect_left(texts, row) - 1
        out: list = []
        while i >= 0 and len(out) < n and texts[i] >= row - 39:
            r = texts[i]
            if out:
                last = out[-1][2]
                between = last - r - 1
                values = bisect.bisect_left(filled, last) - bisect.bisect_right(filled, r)
                if between - values > 2:
                    break
            out.append((cell_name(ws, r, col), cell_text(ws, r, col), r))
            i -= 1
        return [(a, t) for a, t, _r in out]

    def left_texts(ws, row, col):
        """[(cell, text)] in the three cells left of (row, col)."""
        return [(cell_name(ws, row, k), v) for k in range(max(1, col - 3), col) if isinstance(v := cell_text(ws, row, k), str)]

    def list_label(ws, row, col):
        """The label of the column of formulas that (row, col) continues: the texts left of the nearest
        formula above it whose row has some, over formulas only ('Top 3' beside the first of three)."""
        key = ("list-label", id(ws), col)
        memo = cache.setdefault(key, {})
        path, r, found = [], row - 1, []
        while r > 0 and _formula(_value(ws, r, col)):
            if r in memo:
                found = memo[r]
                break
            path.append(r)
            found = left_texts(ws, r, col)
            if found:
                break
            r -= 1
        for q in path:
            memo[q] = found
        return found

    def label_of(ws, row, col):
        """[(cell, text)] labelling the formula at (row, col): the texts in the three cells to its left
        (or, over formulas, the nearest text to its left in the row), else the label of the column of
        formulas it continues, and the two nearest texts above it in its column (a header, a title)."""
        out = left_texts(ws, row, col)
        if not out:
            for k in range(col - 4, max(0, col - 16), -1):
                if isinstance(t := cell_text(ws, row, k), str):
                    out.append((cell_name(ws, row, k), t))
                    break
                if not _formula(_value(ws, row, k)) and not _blank(_value(ws, row, k)):
                    break
        if not out:
            out = list(list_label(ws, row, col))
        return out + texts_above(ws, row, col)

    def span(cols):
        return range(cols[0], min(cols[1], cols[0] + 49) + 1)

    def row_runs(ws, cols):
        """([tops], [(top, bottom)]): the runs of rows holding a value in columns ``cols``."""
        key = ("row-runs", id(ws), cols)
        if key not in cache:
            filled = sorted({c.row for c in _filled(ws) if c.column in span(cols) and not _blank(c.value)})
            runs: list = []
            for r in filled:
                if runs and r == runs[-1][1] + 1:
                    runs[-1][1] = r
                else:
                    runs.append([r, r])
            cache[key] = [r[0] for r in runs], runs
        return cache[key]

    def total_row(ws, r, cols):
        first = next((v for k in span(cols) if not _blank(v := _value(ws, r, k))), None)
        if isinstance(first, str) and not _formula(first) and TOTAL_LABEL.match(first):
            return True
        for k in span(cols):
            f = _formula(_value(ws, r, k))
            m = re.match(r"=\s*([A-Za-z.]+)\s*\(", f or "")
            if m and m.group(1).upper() in AGGREGATES and ":" in f:
                return True
        return False

    def data_rows(ws, idx, cols):
        """({row: position}, number of rows, header rows) of the table in columns ``cols`` that holds row
        ``idx``: its run of rows holding a value, without the header and title rows at its top (rows of
        text above numbers or dates, narrower than the table, or in row 1) and the total rows at its
        bottom (a 'Total' label or an aggregate formula)."""
        tops, runs = row_runs(ws, cols)
        i = bisect.bisect_right(tops, idx) - 1
        if i < 0 or not runs[i][0] <= idx <= runs[i][1]:
            return {}, 0, []
        top, bottom = runs[i]
        key = ("data-rows", id(ws), cols, top)
        if key not in cache:
            def values(r):
                return [v for k in span(cols) if not _blank(v := _value(ws, r, k))]
            width = max(len(values(r)) for r in range(top, min(bottom, top + 50) + 1))
            t = top
            while t < bottom and t - top < 3:
                vs = values(t)
                if not vs or not all(isinstance(v, str) and not _formula(v) for v in vs):
                    break
                typed = any(not isinstance(w := _value(ws, r, k), str) and not _blank(w)
                            for k in span(cols) if isinstance(_value(ws, t, k), str)
                            for r in range(t + 1, min(bottom, t + 20) + 1))
                if not (len(vs) < width or typed or t == 1):
                    break
                t += 1
            b = bottom
            while b > t and bottom - b < 3 and total_row(ws, b, cols):
                b -= 1
            rows = {r: n + 1 for n, r in enumerate(range(t, b + 1))}
            cache[key] = rows, len(rows), list(range(top, t))
        return cache[key]

    def header_of(ws, idx, cols, k):
        """The header text of column ``k`` of the table holding row ``idx`` ('' if none)."""
        heads = data_rows(ws, idx, cols)[2]
        return " ".join(str(v) for r in heads if isinstance(v := cell_text(ws, r, k), str))

    def same_row_refs(ws, r, k, sheet, idx):
        """The columns that the formula at (r, k) reads in row ``idx`` of ``sheet`` as single cells."""
        g = _formula(_value(ws, r, k))
        if not g or str(idx) not in g:
            return []
        return [kc1 for _t, (sh, kc1, kr1, kc2, kr2) in _refs(g, ws.title)
                if sh == sheet and kc1 == kc2 and kr1 == kr2 == idx]

    def numbers(ws, rows, k):
        """{row: number} of column ``k`` when it is a column of numbers (four in five filled cells), else {}."""
        key = ("numbers", id(ws), k, id(rows))
        if key not in cache:
            vals = {r: _value(ws, r, k) for r in rows}
            filled = {r: v for r, v in vals.items() if not _blank(v)}
            nums = {r: v for r, v in filled.items() if isinstance(v, (int, float)) and not isinstance(v, bool)}
            cache[key] = nums if len(nums) >= 2 and len(nums) * 5 >= len(filled) * 4 else {}
        return cache[key]

    def computed_column(ws, rows, k):
        """True when most data rows of column ``k`` hold formulas."""
        key = ("computed", id(ws), k, id(rows))
        if key not in cache:
            cache[key] = sum(1 for r in rows if _formula(_value(ws, r, k))) * 2 > len(rows) > 0
        return cache[key]

    def dates(ws, rows, k, header):
        """{row: sortable date} of column ``k`` when it is a column of dates (dates, 'YYYY-MM-DD' or
        'YYYY/MM/DD' text, or years under a 'Year' header), else {}."""
        key = ("dates", id(ws), k, id(rows))
        if key not in cache:
            out, filled = {}, 0
            for r in rows:
                v = _value(ws, r, k)
                if _blank(v):
                    continue
                filled += 1
                if isinstance(v, (dt.date, dt.datetime)):
                    out[r] = (v.year, v.month, v.day, *((v.hour, v.minute, v.second) if isinstance(v, dt.datetime)
                                                        else (0, 0, 0)))
                elif isinstance(v, str) and (m := DATE_LIKE.match(v)):
                    out[r] = (int(m.group(1)), int(m.group(2)), int(m.group(3)), 0, 0, 0)
                elif isinstance(v, int) and not isinstance(v, bool) and 1900 <= v <= 2100 and YEAR_HEADER.search(header):
                    out[r] = (v, 0, 0, 0, 0, 0)
            cache[key] = out if len(out) >= 2 and len(out) * 5 >= filled * 4 else {}
        return cache[key]

    def stem(w):
        for end in ("ing", "ed", "es", "s"):
            if w.endswith(end) and len(w) - len(end) >= 3:
                return w[: -len(end)]
        return w

    def named(label_text, header):
        """True when the label names the column header ``header``: a word of four or more letters of it
        ('Orders' for 'most ordered'), or, in Chinese, the header itself ('销量')."""
        h = re.sub(r"\(.*?\)", " ", _fold(header))
        t = _fold(label_text)
        if re.search(r"[\u4e00-\u9fff]", h):
            core = "".join(re.findall(r"[\u4e00-\u9fff]+", h))
            return len(core) >= 2 and core in t
        words = {stem(w) for w in re.findall(r"[a-z]+", t)}
        return any(len(w) >= 4 and not _is_position_word(w)
                   and any(s == stem(w) or (len(s) >= 4 and (s.startswith(stem(w)) or stem(w).startswith(s)))
                           for s in words)
                   for w in re.findall(r"[a-z]+", h))

    def value_columns(ws, row, col, a_ws, idx, ref_col, cols, text, sheet):
        """The columns of numbers a word such as 'highest' in the label may be about: the columns whose
        header the label names, else the column the formula reads, else those that formulas beside it
        read in the same row, else (when it reads text) every column of numbers of the table."""
        rows = data_rows(a_ws, idx, cols)[0]
        nums = [k for k in span(cols) if numbers(a_ws, rows, k)]
        computed = [k for k in span(cols) if computed_column(a_ws, rows, k)]
        if ref_col in computed or any(named(text, header_of(a_ws, idx, cols, k)) for k in computed):
            return []  # its numbers are formulas, not known without recalculation
        hit = [k for k in nums if named(text, header_of(a_ws, idx, cols, k))]
        if hit:
            return hit
        if ref_col in nums:
            return [ref_col]
        side = sorted({kc for k in range(max(1, col - 6), col + 7) if k != col
                       for kc in same_row_refs(ws, row, k, sheet, idx) if kc in nums})
        return side or nums

    def holds(claim, ws, row, col, a_ws, idx, ref_col, cols, text, sheet):
        rows, n, _heads = data_rows(a_ws, idx, cols)
        pos = rows.get(idx)
        if pos is None:
            return False
        kind = claim[0]
        if kind in ("pos", "within"):
            at = n - pos + 1 if claim[2] else pos
            return at == claim[1] if kind == "pos" else at <= claim[1]
        if kind == "date":
            for k in span(cols):
                d = dates(a_ws, rows, k, header_of(a_ws, idx, cols, k))
                if idx in d and d[idx] == (max if claim[1] == "max" else min)(d.values()):
                    return True
            return False
        if kind == "edge":
            if pos == (1 if claim[1] == "top" else n):
                return True
            ends: tuple = ("max",) if claim[1] == "top" else ("min",)
        elif kind == "either":
            if pos in (1, n):
                return True
            ends = ("max", "min")
        else:
            ends = (claim[1],)
        k_th = claim[2] if kind == "value" else 1
        for k in value_columns(ws, row, col, a_ws, idx, ref_col, cols, text, sheet):
            nums = numbers(a_ws, rows, k)
            if idx not in nums:
                continue
            for end in ends:
                ranked = sorted(set(nums.values()), reverse=end == "max")
                if len(ranked) >= k_th and nums[idx] == ranked[k_th - 1]:
                    return True
        return False

    def rank_value(ws, row, k):
        """The rank a cell shows: a whole number, a number as text, or =ROW()-n; else None."""
        v = _value(ws, row, k)
        f = _formula(v)
        m = re.fullmatch(r"=\s*ROW\(\s*\)\s*-\s*(\d{1,6})\s*", f or "", re.I)
        if isinstance(v, int) and not isinstance(v, bool):
            return v
        if isinstance(v, str) and not f and v.strip().isascii() and v.strip().isdigit() and len(v.strip()) <= 6:
            return int(v)
        return row - int(m.group(1)) if m else None

    def rank_cells(ws, row, col):
        """[(cell, rank)]: rank numbers left of the formula in its row, in a column of them headed 'Rank',
        '#', 'No.', '名次' or a position ('Top 5')."""
        out = []
        for k in range(max(1, col - 3), col):
            n = rank_value(ws, row, k)
            if n is None:
                continue
            if rank_header(ws, row, k):
                out.append((cell_name(ws, row, k), n))
        return out

    def rank_header(ws, row, k):
        """True when the run of rank numbers holding (row, k) is headed 'Rank', '#', '名次' or a position."""
        memo = cache.setdefault(("rank-header", id(ws), k), {})
        path, r = [], row
        while r > 0 and r not in memo and rank_value(ws, r, k) is not None:
            path.append(r)
            r -= 1
        if r in memo:
            found = memo[r]
        else:
            head = cell_text(ws, r, k) if r > 0 else None
            found = isinstance(head, str) and bool(RANK_HEADER.match(head) or _position_claims(head[:1000]))
        for q in path:
            memo[q] = found
        return found

    def identifying(b_ws, idx, cols):
        """The constants that tell row ``idx`` of the original apart: its label in column A and the
        constants of the table's columns that no other row of the column holds, with their columns."""
        out = [(1, _value(b_ws, idx, 1))]
        for k in span(cols):
            v = _value(b_ws, idx, k)
            if k != 1 and not _blank(v) and not _formula(v) and column_count(b_ws, k, v) == 1:
                out.append((k, v))
        return [(k, v) for k, v in out if not _blank(v) and not isinstance(v, bool) and not _formula(v)]

    def row_words(b_ws, idx, cols):
        """({row: words of its identifying texts}, Counter of the rows each word appears in), for the
        table holding row ``idx``."""
        rows = data_rows(b_ws, idx, cols)[0]
        key = ("row-words", id(b_ws), cols, id(rows))
        if key not in cache:
            per, total = {}, Counter()
            for r in rows:
                words = {stem(x) for _k, v in identifying(b_ws, r, cols) for x in re.findall(r"[^\W\d_]+", _fold(v))}
                per[r] = words
                total.update(words)
            cache[key] = per, total
        return cache[key]

    def names_row(ws, row, col, b_ws, idx, cols, labels_):
        """True when the formula's labels, the cells left of it or the headings above it name the row it
        read before: one of that row's identifying constants (an item name, also in the plural or
        without spaces, a Chinese name inside the label, a number only as '#3' or after its column's
        header, 'Team A'), or a word of them that no other row has ('Kiwi' for 'Kiwi fruit')."""
        own = ws.title == b_ws.title
        context = [str(v) for k in range(1, col) if not (own and cols[0] <= k <= cols[1])
                   and (v := cell_text(ws, row, k)) is not None]
        context += [str(t) for _a, t in labels_]
        if col != 1 and not (own and cols[0] == 1):  # headings in column A (a section 'North' above)
            context += [str(t) for _a, t in texts_above(ws, row, 1)]
        ctx = _fold(" ".join(context))
        compact = re.sub(r"[\W_]+", "", ctx)
        words = {stem(w) for w in re.findall(r"[^\W\d_]+", ctx)}
        per, total = row_words(b_ws, idx, cols)
        own_words = per.get(idx, set())
        for k, v in identifying(b_ws, idx, cols):
            t = _fold(v).strip()
            if isinstance(v, (int, float)) or len(t) < 2:
                head = re.escape(_fold(header_of(b_ws, idx, cols, k)).strip())
                form = r"(?:#|no\.?|n°)\s*" + (rf"|{head}\s*[#:]?\s*" if head else "")
                if re.search(rf"(?<!\w)(?:{form}){re.escape(t)}(?!\w)", ctx):
                    return True
                continue
            if re.search(r"[\u4e00-\u9fff]", t):
                if t in ctx:
                    return True
                continue
            plural = r"(?:s|es)?" if not t.endswith("y") else r"(?:s|es)?|" + re.escape(t[:-1]) + "ies"
            if re.search(r"(?<!\w)(?:" + re.escape(t) + plural + r")(?!\w)", ctx):
                return True
            squeezed = re.sub(r"[\W_]+", "", t)
            if len(squeezed) >= 5 and squeezed in compact:
                return True
            for x in re.findall(r"[^\W\d_]+", t):
                s = stem(x)
                if len(s) >= 3 and s in words and total[s] - (s in own_words) == 0 and not _is_position_word(x):
                    return True
                if len(x) >= 5 and x in compact and total[s] - (s in own_words) == 0 and not _is_position_word(x):
                    return True
        return False

    def reordered(b_ws, a_ws, idx, cols, a_cols, went_to, came_from):
        """True when the rows of the table holding row ``idx`` are in another order: the rows found in
        both files (by their constants, each unique in both) are not in the same order. Rows inserted or
        deleted, with the rest in place, are no reordering. When fewer than two rows can be told apart,
        rows that all moved by one amount (row ``idx``'s old content to ``went_to``, its new content from
        ``came_from``) are none either."""
        b_rows, a_rows = data_rows(b_ws, idx, cols)[0], data_rows(a_ws, idx, a_cols)[0]
        key = ("reordered", id(b_rows), id(a_rows))
        if key not in cache:
            bv, av = row_values(b_ws, cols)[0], row_values(a_ws, a_cols)[0]
            before = [bv.get(r) for r in b_rows]
            after = [av.get(r) for r in a_rows]
            nb, na = Counter(before), Counter(after)
            common = {v for v in nb if v is not None and nb[v] == 1 and na.get(v) == 1}
            order_b = [v for v in before if v in common]
            order_a = [v for v in after if v in common]
            cache[key] = None if len(common) < 2 else order_b != order_a
        found = cache[key]
        return (came_from - idx != idx - went_to) if found is None else found

    def positional(ws, row, col, b_ws, a_ws, ref_sheet, idx, ref_col, went_to, came_from, sibling=True):
        """Why the formula at (row, col), which reads row ``idx`` of ``ref_sheet`` (column ``ref_col``),
        may be meant to read a position in the table rather than one row's content, or None.

        Only after a reordering (reordered): row idx's old content is now in row ``went_to`` and its
        new content was in row ``came_from``. Not when its labels, the cells left of it or the headings
        above it name the row it read. Then: a label names a position (_position_claims) that row
        ``idx`` holds in the edited table, or a rank number beside it is that row's position
        (own_position); or, when it has neither, a formula beside it in its row, or next to it in its
        column, reads the same row and qualifies."""
        if went_to is None or came_from is None or idx in (went_to, came_from):
            return None
        cols = blocks(b_ws).get(ref_col) or (ref_col, ref_col)
        a_cols = blocks(a_ws).get(ref_col) or cols
        if idx not in data_rows(a_ws, idx, a_cols)[0]:
            return None
        if not reordered(b_ws, a_ws, idx, cols, a_cols, went_to, came_from):
            return None
        own = own_position(ws, row, col, b_ws, a_ws, ref_sheet, idx, ref_col)
        if own is not None or not sibling:
            return own or None
        near = [(row, k) for k in range(max(1, col - 3), col + 4) if k != col]
        near += [(r, col) for r in range(max(1, row - 3), row + 4) if r != row]
        for r, k in near:
            for kc in same_row_refs(ws, r, k, ref_sheet, idx):
                if own_position(ws, r, k, b_ws, a_ws, ref_sheet, idx, kc) and not names_row(
                        ws, row, col, b_ws, idx, cols, label_of(ws, row, col)):
                    place = "beside it" if r == row else "next to it in its column"
                    g = _formula(_value(ws, r, k))
                    return f"the formula {place}, {get_column_letter(k)}{r} ={quoted(g[1:])}, reads the same row by its position"
        return None

    def quoted(text):
        text = str(text)
        return text if len(text) <= 60 else text[:57] + "..."

    def own_position(ws, row, col, b_ws, a_ws, ref_sheet, idx, ref_col):
        """Why the formula at (row, col) may be meant to read row ``idx`` by its position, from its own
        labels and rank numbers ('' when they do not show it, None when it has none to go by); computed
        once per cell and row read."""
        key = ("own-position", id(ws), row, col, idx, ref_col)
        if key in cache:
            return cache[key]
        cols = blocks(b_ws).get(ref_col) or (ref_col, ref_col)
        a_cols = blocks(a_ws).get(ref_col) or cols
        labels_ = label_of(ws, row, col)
        field = header_of(a_ws, idx, a_cols, ref_col)
        worded = [(addr, text, found) for addr, text in labels_
                  if (found := _position_claims(str(text)[:1000], str(field)[:1000]))]
        ranks = rank_cells(ws, row, col)
        why = None
        if (worded or ranks) and not names_row(ws, row, col, b_ws, idx, cols, labels_):
            why = ""
            for addr, text, found in worded:
                if any(holds(c, ws, row, col, a_ws, idx, ref_col, a_cols, text, ref_sheet) for c in found):
                    why = f"its label {quoted(text)!r} ({addr}) names a position that row holds now"
                    break
            for addr, n in ranks if not why else ():
                if data_rows(a_ws, idx, a_cols)[0].get(idx) == n:
                    why = f"the rank {n} beside it ({addr}) is that row's position now"
                    break
        cache[key] = why
        return why

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
    neighbours: dict[tuple, dict[str, str]] = {}
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
                # A reference to the formula's own row reads the row it belongs to, unless the formula is a
                # summary beside the table, which does not move with the table's rows.
                own_row = sheet == name and r2 == r1 and r1 == c.row
                across = own_row and beside(b_ws, a_ws, c.row, c.column, c1, c2 if c2 is not None else c1)
                block = blocks(b_ws).get(c1)
                table = block if across else None  # row identity from the table's own columns
                if r2 == r1 and (not own_row or across):
                    other = block is not None and block[0] != 1  # a table beside the one with column A
                    checks.append(("row", r1, c.row, None if other else _value(b_ws, r1, 1),
                                   None if other else _value(a_ws, r1, 1), "label in column A"))
                broken = False
                for axis, idx, own, old, new, part in checks:
                    # A relative reference on the formula's own sheet whose own row/column moved by the
                    # same offset as the one it reads (a filled-down running balance or month-on-month
                    # change after inserted or deleted rows) still reads the same neighbour.
                    neighbour = sheet == name and _relative(text, axis) and not across
                    if old == new:
                        # Same label (a repeated category, or no label): the row's content may still
                        # have moved, as in a sort. Checked for references to another sheet or anchored
                        # ones; a relative one on its own sheet is most likely a filled-down pattern.
                        j = content_moved(b_ws, a_ws, idx, table) if axis == "row" and not neighbour else None
                        if j is not None:
                            was, now = row_values(b_ws, table)[0][idx], row_values(a_ws, table)[0].get(idx)
                            text_ = (f"{f} refers to row {idx} of '{sheet}', whose values {preview(was)} are now in "
                                     f"row {j}; row {idx} now holds " + (preview(now) if now else "no constants"))
                            # Rows reordered (what row idx holds now was in another row before) and a
                            # formula meant to read a position: a question, not a verdict.
                            k = unique_row(b_ws, now, table) if now else None
                            why = positional(after[name], c.row, c.column, b_ws, a_ws, sheet, idx, c1, j, k)
                            if why:
                                findings.append(Finding("warning", "reference-meaning-changed", where,
                                                        f"{text_}: the rows were reordered, and {why}, so it may be "
                                                        "meant to read whatever is in that row now; check that it "
                                                        "should not follow the row it read before"))
                            else:
                                findings.append(Finding("error", "reference-meaning-changed", where, text_))
                            broken = True
                            break
                        continue
                    # The row next to the formula (=D6+C7 in row 7), read without $ on its own sheet by
                    # a formula filled down its column: a running balance or change column. A new row
                    # there, or the row that was next to it before, is the same neighbour; anything
                    # else is a warning, since a sort, or an insert or delete the audit cannot follow,
                    # gives the same picture in a correct edit (a balance refilled after sorting reads
                    # its new neighbour as intended). A one-off formula (VAT =B14*0.2 below a Total)
                    # gets no such leniency.
                    if neighbour and axis == "row" and same_neighbour(b_ws, a_ws, idx, own):
                        continue
                    adjacent = (neighbour and axis == "row" and abs(idx - own) == 1
                                and filled_down(after[name], c.row, c.column, f))
                    if neighbour:
                        if adjacent and (new_row(b_ws, a_ws, idx) or new_row(b_ws, a_ws, own)):
                            continue
                        k = shift_of(b_ws, a_ws, axis, idx)
                        if k and shift_of(b_ws, a_ws, axis, own) == k:
                            continue
                    shift = moved_to(b_ws, a_ws, axis, idx, old, new)
                    if (_blank(new) or shift) and adjacent:
                        # The row was emptied, the cells the formula reads included: a running balance
                        # now starts again from zero there.
                        cleared = emptied(b_ws, a_ws, idx, c1, c2) if _blank(new) and not shift else None
                        if cleared:
                            findings.append(Finding("error", "reference-meaning-changed", where,
                                                    f"{f} refers to row {idx} of '{sheet}', whose label in column A "
                                                    f"was {old!r} and is now empty, and {cleared} and is now "
                                                    "empty too"))
                            broken = True
                            break
                        neighbours.setdefault((name, c.column, bool(shift)), {}).setdefault(
                            where, f"{f} reads row {idx}, whose label in column A was {old!r} and is now "
                                   + (f"{new!r} ({shift})" if shift else "empty"))
                        continue
                    if _blank(new) or shift:
                        what = f"column {get_column_letter(idx)}" if axis == "column" else f"row {idx}"
                        text_ = (f"{f} refers to {what} of '{sheet}', whose {part} was {old!r} and is now "
                                 f"{new!r}" + (f" ({shift})" if shift else ""))
                        # A sort: the old label is now in another row and the new one was in another row
                        # before. A formula meant to read a position (the top row) then reads as intended.
                        why = None
                        if axis == "row" and not _blank(new):
                            went, came = moves(b_ws, a_ws, idx, old, new)
                            why = positional(after[name], c.row, c.column, b_ws, a_ws, sheet, idx, c1, went, came)
                        if why:
                            findings.append(Finding("warning", "reference-meaning-changed", where,
                                                    f"{text_}: the rows were reordered, and {why}, so it may be meant "
                                                    f"to read whatever is in that row now; check that it should not "
                                                    f"follow {old!r}"))
                        else:
                            findings.append(Finding("error", "reference-meaning-changed", where, text_))
                        broken = True
                        break
                    relabelled.setdefault((axis, sheet, idx, old, new, line_stats(b_ws, a_ws, sheet, axis, idx)), []).append(where)
                if broken:
                    break
    for (sheet, col, shifted), found in neighbours.items():
        where, detail = next(iter(found.items()))
        lead = (f"{len(found)} formula(s) filled down column {get_column_letter(col)} of '{sheet}' read the row "
                "next to them")
        if shifted:
            findings.append(Finding("warning", "reference-meaning-changed", where,
                                    f"{lead}, whose label in column A changed (e.g. {detail}): rows were reordered, "
                                    "or inserted or deleted where the audit cannot follow them. That is right for a "
                                    "running balance or change refilled after a sort or insert; check that they should "
                                    "read their new neighbours"))
        else:
            findings.append(Finding("warning", "reference-meaning-changed", where,
                                    f"{lead}, whose label in column A is now empty (e.g. {detail}); check that they "
                                    "should still read that row"))
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
