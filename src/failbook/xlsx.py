"""Small, dependency-light helpers for inspecting .xlsx files in checks."""

from __future__ import annotations

import colorsys
import datetime as _dt
import functools
import math
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any, Iterable, Iterator
from xml.etree import ElementTree as ET

from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

ERROR_VALUES = frozenset(
    {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!"}
)

_REF_RE = re.compile(r"^(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Za-z]{1,3})\$?(\d+)$")


# -- references ---------------------------------------------------------------------
def quote_sheet(name: str) -> str:
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", name) and not re.fullmatch(r"[A-Za-z]{1,3}\d+", name):
        return name
    return "'" + name.replace("'", "''") + "'"


def make_ref(sheet: str, coord: str | tuple[int, int]) -> str:
    if isinstance(coord, tuple):
        row, col = coord
        coord = f"{get_column_letter(col)}{row}"
    return f"{quote_sheet(sheet)}!{coord}"


def split_ref(ref: str) -> tuple[str, str]:
    """``"'Sales 2024'!$C$3"`` -> ``("Sales 2024", "C3")``."""
    m = _REF_RE.match(ref.strip())
    if not m:
        raise ValueError(f"not a single-cell reference: {ref!r}")
    sheet = (m.group(1) or m.group(2)).replace("''", "'")
    return sheet, f"{m.group(3).upper()}{m.group(4)}"


def coord_to_rc(coord: str) -> tuple[int, int]:
    col, row = coordinate_from_string(coord)
    return row, column_index_from_string(col)


# -- values ---------------------------------------------------------------------------
def is_error(value: Any) -> bool:
    return isinstance(value, str) and value.strip().upper() in ERROR_VALUES


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def is_date(value: Any) -> bool:
    return isinstance(value, (_dt.date, _dt.datetime))


def as_date(value: Any) -> _dt.date | None:
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    return None


def norm_text(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).split()).casefold()


def close(a: Any, b: Any, rel: float = 1e-9, abs_tol: float = 1e-6) -> bool:
    if is_number(a) and is_number(b):
        return math.isclose(float(a), float(b), rel_tol=rel, abs_tol=abs_tol)
    if is_date(a) and is_date(b):
        return as_date(a) == as_date(b) and (
            not isinstance(a, _dt.datetime) or not isinstance(b, _dt.datetime) or a == b
        )
    if isinstance(a, str) and isinstance(b, str):
        return a == b
    return a == b


def show(value: Any) -> str:
    """Human readable rendering of a cell value for messages."""
    if value is None:
        return "(empty)"
    if isinstance(value, str):
        return repr(value)
    if isinstance(value, float):
        return f"{value:.6g}"
    if isinstance(value, _dt.datetime) and value.time() == _dt.time(0):
        return value.date().isoformat()
    if is_date(value):
        return value.isoformat()
    return str(value)


def formula_text(cell_value: Any) -> str | None:
    """Return the formula string of a cell loaded with ``data_only=False``."""
    if isinstance(cell_value, ArrayFormula):
        return cell_value.text if str(cell_value.text).startswith("=") else f"={cell_value.text}"
    if isinstance(cell_value, str) and cell_value.startswith("=") and len(cell_value) > 1:
        return cell_value
    return None


# -- sheet navigation -------------------------------------------------------------------
def find_cells(ws, text: str, *, col: int | None = None) -> list[tuple[int, int]]:
    target = norm_text(text)
    hits = []
    for row in ws.iter_rows():
        for cell in row:
            if col is not None and cell.column != col:
                continue
            if isinstance(cell.value, str) and norm_text(cell.value) == target:
                hits.append((cell.row, cell.column))
    return hits


def find_cell(ws, text: str, *, col: int | None = None) -> tuple[int, int] | None:
    hits = find_cells(ws, text, col=col)
    return hits[0] if hits else None


def header_map(ws, row: int = 1) -> dict[str, int]:
    out = {}
    for cell in ws[row]:
        if cell.value is not None and str(cell.value).strip():
            out.setdefault(norm_text(cell.value), cell.column)
    return out


def rows_where(ws, col: int, predicate, start: int = 1) -> Iterator[int]:
    for r in range(start, ws.max_row + 1):
        if predicate(ws.cell(r, col).value):
            yield r


def used_cells(ws) -> Iterator[tuple[int, int, Any]]:
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                yield cell.row, cell.column, cell.value


def number_format(ws, row: int, col: int) -> str:
    return ws.cell(row, col).number_format or "General"


def is_percent_format(fmt: str) -> bool:
    return "%" in (fmt or "")


def ranges_cover(sqref: Any, cells: Iterable[str]) -> bool:
    """Whether a MultiCellRange / sqref string covers every coordinate in ``cells``."""
    from openpyxl.worksheet.cell_range import MultiCellRange

    mcr = sqref if isinstance(sqref, MultiCellRange) else MultiCellRange(str(sqref))
    return all(c in mcr for c in cells)


# -- charts -----------------------------------------------------------------------------
_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def _rels(z: zipfile.ZipFile, part: str) -> dict[str, str]:
    base = posixpath.dirname(part)
    rels_part = posixpath.join(base, "_rels", posixpath.basename(part) + ".rels")
    if rels_part not in z.namelist():
        return {}
    root = ET.fromstring(z.read(rels_part))
    out = {}
    for rel in root.findall("rel:Relationship", _NS):
        target = rel.get("Target", "")
        if rel.get("TargetMode") == "External":
            continue
        if target.startswith("/"):
            resolved = target.lstrip("/")
        else:
            resolved = posixpath.normpath(posixpath.join(base, target))
        out[rel.get("Id")] = resolved
    return out


def charts(path: str | Path) -> list[dict[str, Any]]:
    """List the charts in a workbook with the sheet they are drawn on.

    openpyxl does not load existing charts, so this reads the package directly.
    Each chart is ``{"sheet", "kinds", "title", "series": [{"name", "name_text", "cat", "val"}]}``
    where ``name``/``cat``/``val`` are the formulas the series point at (a range such as
    ``'Monthly'!$B$2:$B$13`` or a defined name such as ``[0]!Revenue``; see
    :func:`resolve_range`) and ``name_text`` is a series name typed in as text.
    """
    found: list[dict[str, Any]] = []
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        wb_root = ET.fromstring(z.read("xl/workbook.xml"))
        wb_rels = _rels(z, "xl/workbook.xml")
        for sheet in wb_root.findall("main:sheets/main:sheet", _NS):
            sheet_name = sheet.get("name")
            sheet_part = wb_rels.get(sheet.get(f"{{{_NS['r']}}}id"))
            if not sheet_part or sheet_part not in names:
                continue
            sheet_rels = _rels(z, sheet_part)
            sroot = ET.fromstring(z.read(sheet_part))
            for drawing in sroot.findall("main:drawing", _NS):
                drawing_part = sheet_rels.get(drawing.get(f"{{{_NS['r']}}}id"))
                if not drawing_part or drawing_part not in names:
                    continue
                for chart_part in _rels(z, drawing_part).values():
                    if chart_part not in names or "/charts/" not in chart_part:
                        continue
                    found.append(_parse_chart(z.read(chart_part), sheet_name))
    return found


def _parse_chart(xml: bytes, sheet_name: str) -> dict[str, Any]:
    root = ET.fromstring(xml)
    plot = root.find("c:chart/c:plotArea", _NS)
    kinds = []
    series = []
    if plot is not None:
        for child in plot:
            tag = child.tag.split("}")[-1]
            if tag.endswith("Chart"):
                kinds.append(tag[: -len("Chart")])
                for ser in child.findall("c:ser", _NS):
                    series.append(
                        {
                            "name": _first_text(ser, "c:tx/c:strRef/c:f"),
                            "name_text": _first_text(ser, "c:tx/c:v"),
                            "cat": _first_text(ser, "c:cat/c:numRef/c:f") or _first_text(ser, "c:cat/c:strRef/c:f")
                            or _first_text(ser, "c:cat/c:multiLvlStrRef/c:f"),
                            "val": _first_text(ser, "c:val/c:numRef/c:f") or _first_text(ser, "c:yVal/c:numRef/c:f"),
                        }
                    )
    title_el = root.find("c:chart/c:title", _NS)
    title = ""
    if title_el is not None:
        title = "".join(t.text or "" for t in title_el.iter(f"{{{_NS['a']}}}t"))
        if not title:
            title = _first_text(title_el, "c:tx/c:strRef/c:f") or ""
    return {"sheet": sheet_name, "kinds": kinds, "title": title, "series": series}


def _first_text(el, path: str) -> str | None:
    found = el.find(path, _NS)
    return found.text if found is not None and found.text else None


_RANGE_RE = re.compile(
    r"^(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$"
)


def parse_range(ref: str) -> tuple[str, int, int, int, int] | None:
    """``"'Monthly'!$B$2:$B$13"`` -> ``("Monthly", min_row, min_col, max_row, max_col)``."""
    if not ref:
        return None
    text = ref.strip()
    if text.startswith("(") and text.endswith(")"):
        text = text[1:-1]
    m = _RANGE_RE.match(text)
    if not m:
        return None
    sheet = (m.group(1) or m.group(2)).replace("''", "'")
    c1, r1 = column_index_from_string(m.group(3).upper()), int(m.group(4))
    c2 = column_index_from_string(m.group(5).upper()) if m.group(5) else c1
    r2 = int(m.group(6)) if m.group(6) else r1
    return sheet, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


_LOCAL_RANGE_RE = re.compile(r"^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$")


#: A defined name: a letter of any script, ``_`` or ``\`` first, then letters, digits, ``_ . \ ?`` (as Excel allows).
_NAME_RE = re.compile(r"(?:[^\W\d]|\\)[\w.\\?]*")


def sheet_named(wb, name: str | None) -> str | None:
    """The workbook's sheet called ``name``, matched case-insensitively as Excel does; None if there is none."""
    if not name:
        return None
    want = name.casefold()
    return next((s for s in wb.sheetnames if s.casefold() == want), None)


def _split_name(text: str, sheet: str | None) -> tuple[str | None, str] | None:
    """``"[0]!Revenue"`` / ``"'My sheet'!Revenue"`` / ``"Revenue"`` -> (scope sheet or None, name), or None."""
    scope, name = sheet, text
    if text.startswith("[0]!"):  # a workbook-level name, as charts write it
        name, scope = text[4:], None
    elif "!" in text:
        sheet_part, name = text.rsplit("!", 1)
        if sheet_part.startswith("'") and sheet_part.endswith("'"):
            sheet_part = sheet_part[1:-1].replace("''", "'")
        if sheet_part.startswith("["):
            return None  # another workbook
        scope = sheet_part
    if not _NAME_RE.fullmatch(name):
        return None
    return scope, name


def _defined_name(wb, name: str, sheet: str | None):
    """The defined name ``name`` as seen from ``sheet``: a sheet-scoped name first, then a workbook name."""
    want = name.casefold()
    scopes = []
    sheet = sheet_named(wb, sheet)
    if sheet:
        scopes.append((sheet, getattr(wb[sheet], "defined_names", {}) or {}))
    scopes.append((None, wb.defined_names))
    for scope, names in scopes:
        for key, defn in list(names.items()):
            if str(key).casefold() == want:
                return scope, defn
    return None


def resolve_range(ref: str | None, wb=None, sheet: str | None = None,
                  _depth: int = 0) -> tuple[str, int, int, int, int] | None:
    """A range formula as ``(sheet, min_row, min_col, max_row, max_col)``, following defined names.

    Accepts ``'Monthly'!$B$2:$B$13``, a range without a sheet (``$H$2:$H$5``, taken to be
    on ``sheet``), and a defined name (``[0]!Revenue``, ``Monthly!Revenue``, ``StatusList``)
    whose definition is such a fixed range, looked up in the openpyxl workbook ``wb``.
    Returns None for anything else (names computed with OFFSET or INDEX, whole columns,
    structured references, unknown names): callers must treat that as "cannot tell",
    never as "no range".
    """
    if not ref:
        return None
    text = str(ref).strip().lstrip("=").strip()
    if text.startswith("(") and text.endswith(")"):
        text = text[1:-1].strip()
    got = parse_range(text)
    if got:
        return got
    if sheet and _LOCAL_RANGE_RE.match(text):
        return parse_range(f"{quote_sheet(sheet)}!{text}")
    if wb is None or _depth > 5:
        return None
    split = _split_name(text, sheet)
    if split is None:
        return None
    scope, name = split
    found = _defined_name(wb, name, scope)
    if found is None:
        return None
    defn_sheet, defn = found
    return resolve_range(getattr(defn, "attr_text", None), wb, defn_sheet, _depth + 1)


# -- colours ----------------------------------------------------------------------------
_THEME_SLOTS = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
                "hlink", "folHlink"]


def _theme_colours(wb) -> dict[int, str]:
    raw = getattr(wb, "loaded_theme", None)
    return _parse_theme(raw) if raw else {}


@functools.lru_cache(maxsize=32)
def _parse_theme(raw: bytes) -> dict[int, str]:
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return {}
    scheme = root.find(".//a:clrScheme", _NS)
    out = {}
    if scheme is None:
        return out
    for i, slot in enumerate(_THEME_SLOTS):
        el = scheme.find(f"a:{slot}", _NS)
        if el is None:
            continue
        srgb = el.find("a:srgbClr", _NS)
        sys_clr = el.find("a:sysClr", _NS)
        if srgb is not None and srgb.get("val"):
            out[i] = srgb.get("val").upper()
        elif sys_clr is not None and sys_clr.get("lastClr"):
            out[i] = sys_clr.get("lastClr").upper()
    return out


def _apply_tint(rgb: str, tint: float) -> str:
    r, g, b = (int(rgb[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + tint) if tint < 0 else l * (1 - tint) + tint
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return "".join(f"{round(v * 255):02X}" for v in (r, g, b))


def colour_rgb(color, wb=None) -> str | None:
    """``RRGGBB`` of an openpyxl Color (rgb, indexed or theme), or None if unset or unknown."""
    if color is None:
        return None
    kind = getattr(color, "type", None)
    if kind == "rgb":
        rgb = str(color.rgb or "")
        if not re.fullmatch(r"[0-9A-Fa-f]{8}|[0-9A-Fa-f]{6}", rgb) or rgb == "00000000":
            return None  # openpyxl's "not set"
        return rgb[-6:].upper()
    if kind == "indexed":
        idx = color.indexed
        if isinstance(idx, int) and 0 <= idx < 64:
            return COLOR_INDEX[idx][-6:].upper()
        return None
    if kind == "theme" and wb is not None:
        base = _theme_colours(wb).get(color.theme)
        if base is None:
            return None
        return _apply_tint(base, float(color.tint or 0.0)) if color.tint else base
    return None


#: Saturation (HSV) from which a red hue counts as red, and below which down to PALE_SATURATION it is a pale red.
RED_SATURATION, PALE_SATURATION = 0.15, 0.05


def is_red(rgb: str | None) -> bool | None:
    """Whether a colour reads as red: a hue from 340 to 15 degrees, HSV saturation at least 15%, value at least 25%.

    Light red fills such as FFC7CE (Excel's "Light Red Fill") count. None if the colour is
    unknown, or if it is a pale red (saturation 5-15%, such as F2DCDB, the theme colour
    "Red, Accent 2, Lighter 80%"), which the checker does not judge either way.
    """
    if rgb is None:
        return None
    r, g, b = (int(rgb[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    hue = h * 360
    if not (hue >= 340 or hue <= 15) or v < 0.25 or s < PALE_SATURATION:
        return False
    return True if s >= RED_SATURATION else None


class Unresolved(str):
    """``"?"``: a colour the checker does not resolve; ``why`` says what it is and why."""

    def __new__(cls, why: str):
        obj = super().__new__(cls, "?")
        obj.why = why
        return obj


def dxf_fill_patterns(path: str | Path) -> list[str | None]:
    """The ``patternType`` of each conditional-format style's fill as written in the file, by dxf index.

    openpyxl reads ``patternType="none"`` as None, the same as an absent attribute, but the
    two differ: no pattern type is how Excel writes a solid conditional fill, ``"none"`` is
    no fill. Entries are None when the attribute (or the fill) is absent.
    """
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            part, rels = "xl/styles.xml", "xl/_rels/workbook.xml.rels"
            for rel in ET.fromstring(z.read(rels)).findall("rel:Relationship", _NS) if rels in names else []:
                if rel.get("Type", "").endswith("/styles") and rel.get("TargetMode") != "External":
                    target = rel.get("Target", "")
                    part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(f"xl/{target}")
            if part not in names:
                return []
            root = ET.fromstring(z.read(part))
    except (OSError, zipfile.BadZipFile, ET.ParseError):
        return []
    out = []
    for dxf in root.findall("main:dxfs/main:dxf", _NS):
        pattern = dxf.find("main:fill/main:patternFill", _NS)
        out.append(pattern.get("patternType") if pattern is not None else None)
    return out


def dxf_colours(dxf, wb=None, raw_pattern: str | None = None) -> dict[str, Any]:
    """The colours a conditional format's style sets: ``{"fill": "RRGGBB" | "?", "font": "RRGGBB" | "?"}``.

    A key is absent when the style does not set that property, and an :class:`Unresolved`
    ``"?"`` when it sets it in a way that is not resolved here (a gradient, an automatic or
    unknown colour). ``raw_pattern`` is the fill's ``patternType`` as written in the file
    (:func:`dxf_fill_patterns`); pass it so that an explicit ``"none"`` reads as no fill.
    Excel writes a conditional format's solid fill as ``bgColor`` with no pattern type, and
    LibreOffice paints that; with no pattern type and only ``fgColor`` LibreOffice paints
    nothing and what Excel paints is not established, so that is "?". With a pattern type,
    ``fgColor`` is used when that is all there is. A fill whose two colours disagree about
    being red is "?".
    """
    out: dict[str, Any] = {}
    if dxf is None:
        return out
    fill = getattr(dxf, "fill", None)
    if fill is not None:
        if fill.__class__.__name__ != "PatternFill":
            out["fill"] = Unresolved("its fill is a gradient")
        elif raw_pattern != "none" and fill.patternType != "none":
            bg, fg = colour_rgb(fill.bgColor, wb), colour_rgb(fill.fgColor, wb)
            set_bg = getattr(fill.bgColor, "rgb", None) != "00000000" or fill.bgColor.type != "rgb"
            set_fg = getattr(fill.fgColor, "rgb", None) != "00000000" or fill.fgColor.type != "rgb"
            if (set_bg and bg is None) or (set_fg and fg is None):
                out["fill"] = Unresolved("its fill colour is automatic or not in the workbook's theme")
            elif bg and fg and is_red(bg) != is_red(fg):
                out["fill"] = Unresolved(f"its fill has a pattern colour {fg} and a background colour {bg} that "
                                         "disagree about being red")
            elif fill.patternType is None and fg and not bg:
                out["fill"] = Unresolved(f"its fill gives only a pattern colour ({fg}) and no pattern type; "
                                         "LibreOffice paints no fill for it, and the checker does not know what "
                                         "Excel paints")
            elif bg or fg:
                out["fill"] = bg or fg
    font = getattr(dxf, "font", None)
    if font is not None and font.color is not None:
        out["font"] = colour_rgb(font.color, wb) or Unresolved("its font colour is automatic or not in the "
                                                               "workbook's theme")
    return out


# -- conditional formats ----------------------------------------------------------------
class Unknown(Exception):
    """A formula or rule uses something the evaluator below does not handle."""


class _Err(str):
    """An Excel error value (``#VALUE!``) inside the evaluator; ``why`` says where it came from, if known."""

    why: str | None = None


def _err(code: str, why: str | None = None) -> _Err:
    e = _Err(code)
    e.why = why
    return e


def _scalar(v: Any, epoch=None) -> Any:
    """A cell value as the evaluator handles it: dates and times become Excel serial numbers."""
    if v is None or isinstance(v, (bool, str)) or is_number(v):
        return v
    if isinstance(v, (_dt.datetime, _dt.date, _dt.time, _dt.timedelta)):
        from openpyxl.utils.datetime import WINDOWS_EPOCH, to_excel

        return float(to_excel(v, epoch or WINDOWS_EPOCH))
    raise Unknown(f"a cell value of type {type(v).__name__} ({v!r}) is not supported")


_FORMULA_REF = re.compile(r"^(?:(?:'((?:[^']|'')+)'|([^'!:]+))!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)$")


def _as_text(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _as_number(v: Any) -> float:
    if v is None:
        return 0.0
    if isinstance(v, bool):
        return float(v)
    if is_number(v):
        return float(v)
    raise Unknown(f"text {v!r} used as a number")


def _compare(a: Any, b: Any, op: str) -> Any:
    if isinstance(a, _Err):
        return a
    if isinstance(b, _Err):
        return b
    # A blank cell compares equal to "" and to 0, like in Excel.
    if a is None:
        a = "" if isinstance(b, str) else (False if isinstance(b, bool) else 0.0)
    if b is None:
        b = "" if isinstance(a, str) else (False if isinstance(a, bool) else 0.0)

    def rank(v):
        if isinstance(v, bool):
            return 2
        if isinstance(v, str):
            return 1
        if is_number(v):
            return 0
        raise Unknown(f"comparing a value of type {type(v).__name__} is not supported")

    ka, kb = rank(a), rank(b)
    if ka != kb:
        a, b = ka, kb
    elif ka == 1:
        a, b = a.casefold(), b.casefold()
    elif ka == 0:
        a, b = float(a), float(b)
    return {"=": a == b, "<>": a != b, "<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[op]


class _Formula:
    """A tiny evaluator for the formulas of simple conditional formats.

    It covers cell references (shifted from the rule's origin cell like Excel does),
    text, numbers, TRUE/FALSE, comparisons, ``&``, ``+ - * /`` and the functions in
    ``_FUNCS``. Anything else raises :class:`Unknown`, so that the caller reports "cannot
    tell" instead of guessing.
    """

    def __init__(self, formula: str, sheet: str, origin: tuple[int, int], target: tuple[int, int], value_at,
                 wb=None):
        self.T = Token
        try:
            items = Tokenizer("=" + str(formula).lstrip("=")).items
        except Exception as exc:
            raise Unknown(f"cannot parse {formula!r}") from exc
        self.toks = [t for t in items if t.type != Token.WSPACE]
        self.i = 0
        self.sheet, self.origin, self.target, self.value_at = sheet, origin, target, value_at
        self.wb = wb

    def run(self) -> Any:
        if not self.toks:
            raise Unknown("empty formula")
        v = self.expr()
        if self.i != len(self.toks):
            raise Unknown(f"unsupported syntax near {self.toks[self.i].value!r}")
        return v

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else None

    def take(self):
        tok = self.peek()
        if tok is None:
            raise Unknown("formula ends unexpectedly")
        self.i += 1
        return tok

    def _infix(self, ops):
        tok = self.peek()
        if tok is not None and tok.type == self.T.OP_IN and tok.value in ops:
            self.i += 1
            return tok.value
        return None

    def expr(self):
        left = self.concat()
        while (op := self._infix({"=", "<>", "<", ">", "<=", ">="})) is not None:
            left = _compare(left, self.concat(), op)
        return left

    def concat(self):
        left = self.additive()
        while self._infix({"&"}) is not None:
            right = self.additive()
            for v in (left, right):
                if isinstance(v, _Err):
                    return v
            left = _as_text(left) + _as_text(right)
        return left

    def additive(self):
        left = self.term()
        while (op := self._infix({"+", "-"})) is not None:
            right = self.term()
            if isinstance(left, _Err) or isinstance(right, _Err):
                left = left if isinstance(left, _Err) else right
                continue
            a, b = _as_number(left), _as_number(right)
            left = a + b if op == "+" else a - b
        return left

    def term(self):
        left = self.unary()
        while (op := self._infix({"*", "/"})) is not None:
            right = self.unary()
            if isinstance(left, _Err) or isinstance(right, _Err):
                left = left if isinstance(left, _Err) else right
                continue
            a, b = _as_number(left), _as_number(right)
            left = a * b if op == "*" else (_Err("#DIV/0!") if b == 0 else a / b)
        return left

    def unary(self):
        tok = self.peek()
        if tok is not None and tok.type == self.T.OP_PRE:
            self.i += 1
            v = self.unary()
            return v if isinstance(v, _Err) else (-_as_number(v) if tok.value == "-" else _as_number(v))
        v = self.primary()
        tok = self.peek()
        if tok is not None and tok.type == self.T.OP_POST and tok.value == "%":
            self.i += 1
            v = v if isinstance(v, _Err) else _as_number(v) / 100
        return v

    def primary(self):
        T = self.T
        tok = self.take()
        if tok.type == T.OPERAND:
            if tok.subtype == T.TEXT:
                return tok.value[1:-1].replace('""', '"')
            if tok.subtype == T.NUMBER:
                return float(tok.value)
            if tok.subtype == T.LOGICAL:
                return tok.value.upper() == "TRUE"
            if tok.subtype == T.ERROR:
                return _Err(tok.value.upper())
            if tok.subtype == T.RANGE:
                return self.ref(tok.value)
        if tok.type == T.PAREN and tok.subtype == T.OPEN:
            v = self.expr()
            close = self.take()
            if close.type != T.PAREN:
                raise Unknown("unbalanced parentheses")
            return v
        if tok.type == T.FUNC and tok.subtype == T.OPEN:
            name = tok.value[:-1].upper().removeprefix("_XLFN.")
            args = []
            nxt = self.peek()
            if nxt is not None and nxt.type == T.FUNC and nxt.subtype == T.CLOSE:
                self.i += 1
            else:
                while True:
                    args.append(self.expr())
                    sep = self.take()
                    if sep.type == T.SEP and sep.subtype == T.ARG:
                        continue
                    if sep.type == T.FUNC and sep.subtype == T.CLOSE:
                        break
                    raise Unknown(f"unsupported syntax near {sep.value!r}")
            func = _FUNCS.get(name)
            if func is None:
                raise Unknown(f"the function {name} is not supported")
            return func(*args)
        raise Unknown(f"unsupported syntax near {tok.value!r}")

    def ref(self, text: str):
        m = _FORMULA_REF.match(text.lstrip("@"))
        if not m:
            return self.name(text)
        sheet = (m.group(1) or "").replace("''", "'") or m.group(2) or self.sheet
        col = column_index_from_string(m.group(4).upper())
        row = int(m.group(6))
        if not m.group(5):
            row += self.target[0] - self.origin[0]
        if not m.group(3):
            col += self.target[1] - self.origin[1]
        if row < 1 or col < 1:
            return _Err("#REF!")
        v = self.value_at(sheet, row, col)
        if isinstance(v, str) and is_error(v):
            return _Err(v.strip().upper())
        return _scalar(v, getattr(self.wb, "epoch", None))

    def name(self, text: str):
        """A defined name: ``#NAME?`` if the workbook does not define it, as in Excel."""
        split = _split_name(text.lstrip("@"), self.sheet)
        if split is None:
            raise Unknown(f"the reference {text!r} is not a single cell, which is not supported")
        if self.wb is None:
            raise Unknown(f"{text!r} is a name, which is not looked up without the workbook")
        scope, name = split
        if scope is not None and sheet_named(self.wb, scope) is None:
            raise Unknown(f"the sheet {scope!r} in {text!r} does not exist")
        found = _defined_name(self.wb, name, scope)
        if found is None:
            tables = {str(t).casefold() for ws in self.wb.worksheets for t in getattr(ws, "tables", {})}
            if name.casefold() in tables:
                raise Unknown(f"the table name {name!r} is not evaluated by the checker")
            return _err("#NAME?", f"the name {name!r} is not defined in the workbook")
        defn = getattr(found[1], "attr_text", None)
        raise Unknown(f"the defined name {name!r} (= {defn}) is not evaluated by the checker")


def _args(n_min, n_max):
    def deco(fn):
        def wrapped(*args):
            if not n_min <= len(args) <= n_max:
                raise Unknown(f"a function called with {len(args)} arguments is not supported")
            return fn(*args)
        return wrapped
    return deco


def _first_err(*vals):
    for v in vals:
        if isinstance(v, _Err):
            return v
    return None


def _bool(v):
    if isinstance(v, _Err):
        return v
    if isinstance(v, bool):
        return v
    if v is None:
        return False
    if is_number(v):
        return v != 0
    raise Unknown(f"text {v!r} used as TRUE/FALSE")


def _search(case_sensitive):
    def fn(find, within, start=1.0):
        err = _first_err(find, within, start)
        if err:
            return err
        find, within = _as_text(find), _as_text(within)
        if not case_sensitive and re.search(r"[?*~]", find):
            raise Unknown("SEARCH with wildcards")
        pos = (within if case_sensitive else within.casefold()).find(
            find if case_sensitive else find.casefold(), int(_as_number(start)) - 1)
        return float(pos + 1) if pos >= 0 else _Err("#VALUE!")
    return _args(2, 3)(fn)


def _text_fn(op):
    def fn(*args):
        err = _first_err(*args)
        return err if err else op(*args)
    return fn


_FUNCS = {
    "AND": lambda *a: _first_err(*[_bool(x) for x in a]) or all(_bool(x) for x in a),
    "OR": lambda *a: _first_err(*[_bool(x) for x in a]) or any(_bool(x) for x in a),
    "NOT": _args(1, 1)(lambda a: _first_err(_bool(a)) or not _bool(a)),
    "IF": _args(2, 3)(lambda c, a, b=False: _first_err(_bool(c)) or (a if _bool(c) else b)),
    "EXACT": _args(2, 2)(_text_fn(lambda a, b: _as_text(a) == _as_text(b))),
    "SEARCH": _search(False),
    "FIND": _search(True),
    "ISNUMBER": _args(1, 1)(lambda a: is_number(a) and not isinstance(a, _Err)),
    "ISTEXT": _args(1, 1)(lambda a: isinstance(a, str) and not isinstance(a, _Err)),
    "ISERROR": _args(1, 1)(lambda a: isinstance(a, _Err)),
    "ISERR": _args(1, 1)(lambda a: isinstance(a, _Err) and a != "#N/A"),
    "ISBLANK": _args(1, 1)(lambda a: a is None),
    "LEN": _args(1, 1)(_text_fn(lambda a: float(len(_as_text(a))))),
    "TRIM": _args(1, 1)(_text_fn(lambda a: " ".join(p for p in _as_text(a).split(" ") if p))),
    "UPPER": _args(1, 1)(_text_fn(lambda a: _as_text(a).upper())),
    "LOWER": _args(1, 1)(_text_fn(lambda a: _as_text(a).lower())),
    "LEFT": _args(1, 2)(_text_fn(lambda a, n=1.0: _as_text(a)[:max(0, int(_as_number(n)))])),
    "RIGHT": _args(1, 2)(_text_fn(lambda a, n=1.0: _as_text(a)[::-1][:max(0, int(_as_number(n)))][::-1])),
}


def eval_cf_formula(formula: str, sheet: str, origin: tuple[int, int], target: tuple[int, int], value_at,
                    wb=None) -> Any:
    """Evaluate a conditional-format formula for the cell ``target`` (row, col).

    Relative references are shifted from ``origin``, the top-left cell of the rule's
    range, as Excel and LibreOffice do. ``value_at(sheet, row, col)`` supplies cell
    values; dates in them are compared as Excel serial numbers. With the openpyxl
    workbook ``wb``, a name the workbook does not define gives ``#NAME?``. Raises
    :class:`Unknown` for formulas outside the supported subset.
    """
    return _Formula(formula, sheet, origin, target, value_at, wb).run()


_CELL_IS = {"equal": "=", "notEqual": "<>", "greaterThan": ">", "lessThan": "<",
            "greaterThanOrEqual": ">=", "lessThanOrEqual": "<="}
_TEXT_RULES = {
    "containsText": lambda v, t: t.casefold() in v.casefold(),
    "notContainsText": lambda v, t: t.casefold() not in v.casefold(),
    "beginsWith": lambda v, t: v.casefold().startswith(t.casefold()),
    "endsWith": lambda v, t: v.casefold().endswith(t.casefold()),
}
#: Rule types that never colour a text cell through a differential style.
_NO_DXF_RULES = {"colorScale", "dataBar", "iconSet"}


def _truthy(v: Any) -> bool:
    if isinstance(v, _Err):
        return False  # a formula that returns an error does not apply its format
    if isinstance(v, bool):
        return v
    if v is None:
        return False
    if is_number(v):
        return v != 0
    raise Unknown(f"the formula returns text ({v!r}), not TRUE or FALSE")


def cf_rule_applies(cf_rule, sheet: str, origin: tuple[int, int], target: tuple[int, int], value_at,
                    wb=None) -> bool:
    """Whether an openpyxl conditional-format rule applies its style to the cell ``target``.

    Handles ``cellIs``, ``expression``, the text rules (contains / does not contain /
    begins with / ends with), blanks and errors. Raises :class:`Unknown` for other rule
    types (top 10, above average, duplicates, dates ...) and for formulas outside the
    subset :func:`eval_cf_formula` supports.
    """
    return cf_rule_outcome(cf_rule, sheet, origin, target, value_at, wb)[0]


def cf_rule_outcome(cf_rule, sheet: str, origin: tuple[int, int], target: tuple[int, int], value_at,
                    wb=None) -> tuple[bool, str | None]:
    """:func:`cf_rule_applies`, plus a note when the rule does not apply because it gives an
    error value there (Excel and LibreOffice then do not apply the format); otherwise None."""
    errors: list[_Err] = []

    def ev(f):
        return eval_cf_formula(f, sheet, origin, target, value_at, wb)

    def truthy(v):
        if isinstance(v, _Err):
            errors.append(v)
        return _truthy(v)

    if cf_rule.type in _NO_DXF_RULES:
        return False, None  # colour scales, data bars and icon sets apply no differential style
    value = value_at(sheet, *target)
    if isinstance(value, str) and is_error(value):
        value = _Err(value.strip().upper())
    else:
        value = _scalar(value, getattr(wb, "epoch", None))
    applies = _rule_applies(cf_rule, value, ev, truthy)
    if applies or not errors:
        return applies, None
    err = errors[0]
    return False, f"gives {err} there{f' ({err.why})' if err.why else ''}, so it does not apply"


def _rule_applies(cf_rule, value, ev, truthy) -> bool:
    kind = cf_rule.type
    formulas = [str(f) for f in (cf_rule.formula or [])]
    if kind == "cellIs":
        op = cf_rule.operator
        if op in _CELL_IS and formulas:
            return truthy(_compare(value, ev(formulas[0]), _CELL_IS[op]))
        if op in ("between", "notBetween") and len(formulas) >= 2:
            lo, hi = ev(formulas[0]), ev(formulas[1])
            err = _first_err(value, lo, hi)
            if err:
                return truthy(err)  # an error applies neither "between" nor "not between"
            inside = (truthy(_compare(value, lo, ">=")) and truthy(_compare(value, hi, "<="))) or (
                truthy(_compare(value, hi, ">=")) and truthy(_compare(value, lo, "<=")))
            return inside if op == "between" else not inside
        raise Unknown(f"the cell-value operator {op!r} is not supported")
    if kind == "expression":
        if not formulas:
            raise Unknown("it is a formula rule without a formula")
        return truthy(ev(formulas[0]))
    if kind in _TEXT_RULES:
        if not formulas:
            raise Unknown(f"it is a {kind} rule without a formula")
        by_formula = truthy(ev(formulas[0]))
        text = getattr(cf_rule, "text", None)
        if text not in (None, "") and not isinstance(value, _Err):
            by_text = _TEXT_RULES[kind](_as_text(value), str(text))
            if by_text != by_formula:
                raise Unknown(f"its text {text!r} and its formula {formulas[0]!r} disagree for this cell")
        return by_formula
    if kind in ("containsBlanks", "notContainsBlanks", "containsErrors", "notContainsErrors"):
        if formulas:
            return truthy(ev(formulas[0]))
        if kind.endswith("Blanks"):
            blank = _as_text(value).strip() == ""
            return blank if kind == "containsBlanks" else not blank
        err = is_error(value)
        return err if kind == "containsErrors" else not err
    raise Unknown(f"rules of type {kind!r} are not supported")
