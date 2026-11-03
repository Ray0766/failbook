"""Helpers for writing cases: deterministic fake data and common workbook styling.

All data is fictional. Names, companies and products are generated from the
small lists below; any resemblance to real entities is coincidental.
"""

from __future__ import annotations

import datetime as dt
import random
from typing import Any, Iterable, Sequence

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

FIRST_NAMES = [
    "Ava", "Ben", "Chen", "Dina", "Eli", "Fatima", "Gao", "Hana", "Ivan", "Jia", "Kofi", "Lena",
    "Mateo", "Nora", "Omar", "Priya", "Quinn", "Rosa", "Sven", "Tara", "Umar", "Vera", "Wei", "Xin",
    "Yuki", "Zoe", "Arjun", "Bea", "Carlos", "Dmitri", "Emma", "Femi",
]
LAST_NAMES = [
    "Adams", "Bauer", "Costa", "Diaz", "Evans", "Fischer", "Garcia", "Huang", "Ito", "Jensen",
    "Khan", "Li", "Moreau", "Novak", "Okafor", "Park", "Quist", "Rossi", "Sato", "Tan", "Ueda",
    "Varga", "Wang", "Xu", "Yilmaz", "Zhou",
]
COMPANIES = [
    "Acme Tools", "Blue Harbor", "Cedar & Co", "Delta Foods", "Evergreen Labs", "Foxglove Ltd",
    "Granite Works", "Helix Media", "Ironwood", "Juniper Retail", "Kestrel Air", "Lumen Studio",
    "Maple Logistics", "Nimbus Cloud", "Orchid Health", "Pioneer Farms", "Quartz Optics",
    "Redwood Bikes", "Summit Gear", "Tidewater Inc",
]
PRODUCTS = [
    "Desk Lamp", "Office Chair", "Notebook A5", "Gel Pen (10)", "Monitor Arm", "USB-C Hub",
    "Webcam HD", "Mouse Pad", "Paper A4 (500)", "Stapler", "Whiteboard", "Label Printer",
    "Headset", "Keyboard", "Cable Tray", "Desk Mat", "Laptop Stand", "Sticky Notes",
    "File Cabinet", "Bookshelf", "Shredder", "Coffee Maker", "Water Filter", "Plant Pot",
]
REGIONS = ["North", "South", "East", "West", "Central"]
DEPARTMENTS = ["Engineering", "Finance", "Marketing", "Operations", "Sales", "Support"]

HEADER_FILL = PatternFill("solid", start_color="FF1F4E79", end_color="FF1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFFFF")
TOTAL_FONT = Font(bold=True)
THIN = Side(style="thin", color="FFBFBFBF")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

DATE_FMT = "yyyy-mm-dd"
MONEY_FMT = "#,##0.00"
INT_FMT = "#,##0"
PCT_FMT = "0.0%"


def rng(seed: int, salt: str) -> random.Random:
    """A random generator unique to (case, seed); seed 0 is the canonical variant."""
    return random.Random(f"failbook:{salt}:{seed}")


def people(r: random.Random, n: int) -> list[str]:
    out: list[str] = []
    while len(out) < n:
        name = f"{r.choice(FIRST_NAMES)} {r.choice(LAST_NAMES)}"
        if name not in out:
            out.append(name)
    return out


def sample(r: random.Random, items: Sequence[Any], n: int) -> list[Any]:
    return r.sample(list(items), n)


def style_header(ws, row: int = 1, ncols: int | None = None) -> None:
    ncols = ncols or ws.max_column
    for col in range(1, ncols + 1):
        cell = ws.cell(row, col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BOX


def set_widths(ws, widths: dict[str, float]) -> None:
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def md_table(headers: Sequence[str], rows: Iterable[Sequence[Any]]) -> str:
    def fmt(v: Any) -> str:
        if isinstance(v, float):
            return f"{v:,.2f}" if abs(v) >= 1000 else (f"{v:.2f}" if v != int(v) else f"{int(v)}")
        if isinstance(v, (dt.date, dt.datetime)):
            return v.isoformat()
        return str(v)

    lines = ["| " + " | ".join(headers) + " |", "|" + "|".join("---" for _ in headers) + "|"]
    for row in rows:
        lines.append("| " + " | ".join(fmt(v) for v in row) + " |")
    return "\n".join(lines)


def iso(d: dt.date) -> str:
    return d.isoformat()


def from_iso(s: str) -> dt.date:
    return dt.date.fromisoformat(s)
