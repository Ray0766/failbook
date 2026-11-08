"""Regenerating the committed artifacts of a case and checking they are current."""

from __future__ import annotations

import datetime as dt
import json
import shutil
import tempfile
from pathlib import Path
from typing import Any

import openpyxl

from .case import Case
from .engine import default_engine


def build_case(case: Case, *, engine=None) -> None:
    """Rebuild ``input/`` (seed 0), ``failure/`` and ``failure/grade.json``."""
    from .grade import grade

    engine = engine or default_engine()
    for d in (case.input_dir, case.failure_dir):
        if d.exists():
            for child in d.iterdir():
                if child.is_file() and child.name != "README.md":
                    child.unlink()
        d.mkdir(parents=True, exist_ok=True)
    meta = case.build(0, case.input_dir, engine=engine)
    with tempfile.TemporaryDirectory(prefix="failbook-build-failure-") as tmp:
        work = Path(tmp) / "work"
        shutil.copytree(case.input_dir, work)
        case.solve("failure", work, meta)
        for name in case.outputs:
            shutil.copyfile(work / name, case.failure_dir / name)
    report = grade(case, case.failure_dir, 0, engine=engine, input_dir=case.input_dir)
    summary = {
        "case": case.id,
        "case_version": case.version,
        "seed": 0,
        "verdict": report.verdict,
        "rules": [r.to_dict() for r in report.rules],
    }
    (case.failure_dir / "grade.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def signature(path: Path) -> dict[str, Any]:
    """Content of a workbook that must not drift: sheet names, values and formulas."""
    wb = openpyxl.load_workbook(path)
    out: dict[str, Any] = {"sheets": wb.sheetnames, "cells": {}}
    for ws in wb.worksheets:
        cells = {}
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if v is None:
                    continue
                if isinstance(v, float):
                    v = round(v, 9)
                elif isinstance(v, (dt.date, dt.datetime)):
                    v = v.isoformat()
                elif not isinstance(v, (int, str, bool)):
                    v = str(getattr(v, "text", v))
                cells[c.coordinate] = v
        out["cells"][ws.title] = cells
    return out


def _diff(a: dict, b: dict) -> str | None:
    if a["sheets"] != b["sheets"]:
        return f"sheets differ: {a['sheets']} vs {b['sheets']}"
    for sheet in a["sheets"]:
        ca, cb = a["cells"][sheet], b["cells"][sheet]
        for coord in sorted(set(ca) | set(cb)):
            if ca.get(coord) != cb.get(coord):
                return f"{sheet}!{coord}: committed {ca.get(coord)!r} vs rebuilt {cb.get(coord)!r}"
    return None


def check_committed(case: Case, *, engine=None) -> list[str]:
    """Compare committed inputs and failure artifacts with a fresh rebuild."""
    engine = engine or default_engine()
    problems = []
    with tempfile.TemporaryDirectory(prefix="failbook-check-") as tmp:
        tmp = Path(tmp)
        meta = case.build(0, tmp / "input", engine=engine)
        for name in case.inputs:
            committed = case.input_dir / name
            if not committed.is_file():
                problems.append(f"input/{name} is missing (run `failbook build {case.id}`)")
                continue
            d = _diff(signature(committed), signature(tmp / "input" / name))
            if d:
                problems.append(f"input/{name} is stale: {d}")
        work = tmp / "work"
        shutil.copytree(case.input_dir if case.input_dir.is_dir() else tmp / "input", work)
        case.solve("failure", work, meta)
        for name in case.outputs:
            committed = case.failure_dir / name
            if not committed.is_file():
                problems.append(f"failure/{name} is missing (run `failbook build {case.id}`)")
                continue
            d = _diff(signature(committed), signature(work / name))
            if d:
                problems.append(f"failure/{name} is stale: {d}")
        if not (case.failure_dir / "grade.json").is_file():
            problems.append("failure/grade.json is missing")
    return problems
