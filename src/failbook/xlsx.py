"""Small, dependency-light helpers for inspecting .xlsx files in checks."""

from __future__ import annotations

import datetime as _dt
import math
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any, Iterable, Iterator
from xml.etree import ElementTree as ET

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
    Each chart is ``{"sheet", "kinds", "title", "series": [{"name", "cat", "val"}]}``
    where ``cat``/``val`` are the range formulas the series point at.
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
                            "cat": _first_text(ser, "c:cat/c:numRef/c:f") or _first_text(ser, "c:cat/c:strRef/c:f"),
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
