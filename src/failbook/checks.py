"""The rule API that case checkers are written against, and the grader that runs them.

A rule is a plain function decorated with :func:`rule`. It receives a
:class:`Context` and either returns (pass), raises :class:`Fail` (or uses
``assert``) when the work is wrong, or raises :class:`Inconclusive` when the
checker cannot decide (for example because the recalculation engine does not
support a function the agent used). Inconclusive is never counted as a failure
of the agent.
"""

from __future__ import annotations

import datetime as _dt
import re
import shutil
import tempfile
import traceback
from collections import deque
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Any, Callable, Iterable

import openpyxl
from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.utils.cell import range_boundaries

from . import xlsx as X
from .engine import default_engine

if TYPE_CHECKING:  # pragma: no cover
    from .case import Case

PASS, FAIL, INCONCLUSIVE = "pass", "fail", "inconclusive"

#: What a rule looks at. Used to report results by kind and for ablations
#: (for example "would a check of the values alone have caught this?").
RULE_KINDS = {
    "content": "the requested change is present, with the right values and types",
    "value": "recalculated results are correct",
    "live": "results follow when an input changes (formulas, not pasted numbers)",
    "preserve": "specific things the task did not ask to change (named in each rule) are unchanged",
    "structure": "formats, charts, validation and other non-cell structure",
}


class Fail(Exception):
    """The agent's output does not meet a requirement."""


class Inconclusive(Exception):
    """The checker cannot decide; never counted against the agent."""


def require(condition: Any, message: str) -> None:
    if not condition:
        raise Fail(message)


@dataclass
class RuleSpec:
    id: str
    title: str
    requirement: str
    title_zh: str | None
    func: Callable[["Context"], Any]
    kind: str = "value"


def rule(id: str, title: str, *, requirement: str, kind: str = "value", title_zh: str | None = None):
    """Register a check.

    ``requirement`` quotes (or closely paraphrases) the sentence of the task that
    this rule verifies. Every rule must trace back to something the agent was
    actually asked to do; see "Rules for rules" in docs/WRITING_CASES.md.
    """

    if kind not in RULE_KINDS:
        raise ValueError(f"rule {id}: kind must be one of {sorted(RULE_KINDS)}")

    def deco(fn):
        fn.__failbook_rule__ = RuleSpec(id=id, title=title, requirement=requirement, title_zh=title_zh,
                                        func=fn, kind=kind)
        return fn

    return deco


@dataclass
class RuleResult:
    id: str
    title: str
    status: str
    detail: str = ""
    kind: str = "value"

    def to_dict(self) -> dict[str, Any]:
        return {"id": self.id, "kind": self.kind, "title": self.title, "status": self.status, "detail": self.detail}


@dataclass
class GradeReport:
    case: str
    case_version: int
    seed: int
    verdict: str
    rules: list[RuleResult]
    engine: str
    failbook_version: str
    graded_at: str = field(default_factory=lambda: _dt.datetime.now(_dt.timezone.utc).isoformat(timespec="seconds"))

    @property
    def score(self) -> float:
        decided = [r for r in self.rules if r.status != INCONCLUSIVE]
        if not decided:
            return 0.0
        return sum(r.status == PASS for r in decided) / len(decided)

    def to_dict(self) -> dict[str, Any]:
        return {
            "case": self.case,
            "case_version": self.case_version,
            "seed": self.seed,
            "verdict": self.verdict,
            "score": round(self.score, 4),
            "rules": [r.to_dict() for r in self.rules],
            "engine": self.engine,
            "failbook_version": self.failbook_version,
            "graded_at": self.graded_at,
        }


class Values:
    """Recalculated cell values of one workbook (one scenario)."""

    def __init__(self, path: Path, label: str, source: str):
        self.path = path
        self.label = label
        self.source = source  # the output file these values were computed from
        self.wb = openpyxl.load_workbook(path, data_only=True)

    def __getitem__(self, ref: str) -> Any:
        sheet, coord = X.split_ref(ref)
        if sheet not in self.wb.sheetnames:
            raise Fail(f"sheet {sheet!r} not found (sheets: {', '.join(self.wb.sheetnames)})")
        return self.wb[sheet][coord].value

    def ws(self, sheet: str):
        if sheet not in self.wb.sheetnames:
            raise Fail(f"sheet {sheet!r} not found (sheets: {', '.join(self.wb.sheetnames)})")
        return self.wb[sheet]


_CELLISH = re.compile(r"^\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$|^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$|^\$?\d+:\$?\d+$")

#: How many cells the #NAME? diagnosis may visit while following an error to its origin.
_PRECEDENT_BUDGET = 50_000


def _is_name_error(value: Any) -> bool:
    return isinstance(value, str) and value.strip().upper() == "#NAME?"


def _split_sheet(text: str) -> tuple[list[str], str] | None:
    """``"'Jan 1:Mar 1'!B2"`` -> ``(["Jan 1", "Mar 1"], "B2")``; None if there is no sheet part.

    A 3-D span gives both end sheets (sheet names cannot contain ``:``). An
    external reference (``[1]Sheet!A1``) gives an empty sheet list.
    """
    if text.startswith("'"):
        i = 1
        while i < len(text):
            if text[i] == "'":
                if text[i + 1:i + 2] == "'":
                    i += 2
                    continue
                break
            i += 1
        if text[i + 1:i + 2] != "!":
            return None
        name, rest = text[1:i].replace("''", "'"), text[i + 2:]
    else:
        if "!" not in text:
            return None
        name, rest = text.split("!", 1)
    if name.startswith("[") or not name:
        return [], rest
    return name.split(":"), rest


class _References:
    """What the sheets, defined names and tables of one workbook make of a formula."""

    def __init__(self, wb):
        self.wb = wb
        self.sheets = {s.casefold(): s for s in wb.sheetnames}
        self.global_names = {n.casefold(): d for n, d in _items(wb.defined_names)}
        self.local_names = {ws.title: {n.casefold(): d for n, d in _items(getattr(ws, "defined_names", {}))}
                            for ws in wb.worksheets}
        self.any_name = set(self.global_names).union(*self.local_names.values())
        self.tables = {}
        for ws in wb.worksheets:
            for t_name, table in _items(getattr(ws, "tables", {})):
                ref = table if isinstance(table, str) else getattr(table, "ref", None)
                if ref:
                    self.tables[t_name.casefold()] = (ws.title, ref)

    # -- what a name or table refers to ------------------------------------------
    def _name(self, sheet: str, name: str):
        local = self.local_names.get(sheet, {})
        return local.get(name.casefold()) or self.global_names.get(name.casefold())

    def _name_problem(self, defn, seen: frozenset = frozenset()) -> str | None:
        text = str(getattr(defn, "attr_text", "") or "")
        if defn.name.casefold() in seen:
            return None
        seen = seen | {defn.name.casefold()}
        tokens = _tokens("=" + text)
        if tokens is None:
            return None
        scope = self.wb.sheetnames[defn.localSheetId] if getattr(defn, "localSheetId", None) is not None \
            and defn.localSheetId < len(self.wb.sheetnames) else None
        for tok in tokens:
            if _is_ref_error(tok):
                return "contains a broken reference (#REF!)"
            if tok.type != Token.OPERAND or tok.subtype != Token.RANGE:
                continue
            split = _split_sheet(tok.value)
            if split is not None:
                missing = [s for s in split[0] if s.casefold() not in self.sheets]
                if missing:
                    return f"points to a sheet that does not exist ({missing[0]!r})"
            elif not _CELLISH.match(tok.value) and "[" not in tok.value:
                inner = self._name(scope, tok.value) if scope else self.global_names.get(tok.value.casefold())
                if inner is not None:
                    problem = self._name_problem(inner, seen)
                    if problem:
                        return f"{problem} (through the name {inner.name} = {inner.attr_text})"
        return None

    # -- one formula ------------------------------------------------------------------
    def problem(self, sheet: str, formula: str) -> str | None:
        """A fact about ``formula`` (on ``sheet``) that makes it a broken reference, or None."""
        tokens = _tokens(formula)
        if tokens is None:
            return None
        local = _local_names(tokens)
        for tok in tokens:
            if _is_ref_error(tok):
                return "contains a broken reference"
            if tok.type != Token.OPERAND or tok.subtype != Token.RANGE:
                continue
            text = tok.value
            split = _split_sheet(text)
            if split is not None:
                sheets, rest = split
                for s in sheets:
                    if s.casefold() not in self.sheets:
                        return f"refers to a sheet that does not exist ({s!r})"
                if len(sheets) == 1 and rest and not _CELLISH.match(rest) and "!" not in rest:
                    defn = self.local_names.get(self.sheets[sheets[0].casefold()], {}).get(rest.casefold())
                    if defn is not None:
                        problem = self._name_problem(defn)
                        if problem:
                            return f"uses the name {text} (= {defn.attr_text}), which {problem}"
                tail = rest.split(":", 1)[1] if ":" in rest else ""
                tail_split = _split_sheet(tail) if "!" in tail else None
                if tail_split:
                    for s in tail_split[0]:
                        if s.casefold() not in self.sheets:
                            return f"refers to a sheet that does not exist ({s!r})"
                continue
            if "[" in text:  # structured reference: Table[Column], [@Column], ...
                table = text.split("[", 1)[0]
                if table and table.casefold() not in self.tables and table.casefold() not in self.any_name:
                    return f"refers to a table that does not exist ({table!r})"
                continue
            if _CELLISH.match(text) or text.casefold() in local or text.lower().startswith("_xl"):
                continue  # a cell, a name the formula defines itself, or an Excel-internal token (_xlpm., _xleta.)
            defn = self._name(sheet, text)
            if defn is not None:
                problem = self._name_problem(defn)
                if problem:
                    return f"uses the name {text} (= {defn.attr_text}), which {problem}"
                continue
            if text.casefold() in self.any_name or text.casefold() in self.tables:
                continue
            return f"refers to an undefined name {text!r}"
        return None

    def precedents(self, sheet: str, formula: str) -> Iterable[tuple]:
        """Cell ranges ``formula`` reads directly: (sheet, min_col, min_row, max_col, max_row).

        A bound is None for a whole column or row.
        """
        tokens = _tokens(formula) or []
        for tok in tokens:
            if tok.type != Token.OPERAND or tok.subtype != Token.RANGE:
                continue
            yield from self._ranges(sheet, tok.value, frozenset())

    def _ranges(self, sheet: str, text: str, seen: frozenset):
        split = _split_sheet(text)
        if split is not None:
            sheets, rest = split
            if not sheets or not _CELLISH.match(rest):
                return
            titles = [self.sheets.get(s.casefold()) for s in sheets]
            if None in titles:
                return
            order = self.wb.sheetnames
            lo, hi = sorted(order.index(t) for t in (titles[0], titles[-1]))
            for title in order[lo:hi + 1]:
                yield (title, *_bounds(rest))
            return
        if _CELLISH.match(text):
            yield (sheet, *_bounds(text))
            return
        if "[" in text:
            table = self.tables.get(text.split("[", 1)[0].casefold())
            if table:
                yield (table[0], *_bounds(table[1]))
            return
        if text.casefold() in self.tables and self._name(sheet, text) is None:
            title, ref = self.tables[text.casefold()]
            yield (title, *_bounds(ref))
            return
        defn = self._name(sheet, text)
        if defn is None or defn.name.casefold() in seen:
            return
        for tok in _tokens("=" + str(defn.attr_text or "")) or []:
            if tok.type == Token.OPERAND and tok.subtype == Token.RANGE:
                yield from self._ranges(sheet, tok.value, seen | {defn.name.casefold()})


def _items(mapping) -> list:
    try:
        return list(mapping.items())
    except Exception:  # pragma: no cover - unexpected openpyxl container
        return []


def _tokens(formula: str) -> list[Token] | None:
    """The formula's tokens without white space, or None if it cannot be parsed.

    A range operand loses the ``@`` (implicit intersection) and the ``:`` it can
    carry when the range operator joins it to a function, as in ``OFFSET(...):A5``.
    """
    try:
        items = Tokenizer(formula).items
    except Exception:
        return None
    out = []
    for tok in items:
        if tok.type == Token.WSPACE:
            continue
        if tok.type == Token.OPERAND and tok.subtype == Token.RANGE:
            value = tok.value.lstrip("@").strip(":")
            if not value:
                continue
            tok = Token(value, Token.OPERAND, Token.RANGE)
        out.append(tok)
    return out


def _is_ref_error(tok) -> bool:
    return tok.type == Token.OPERAND and tok.subtype in (Token.ERROR, Token.RANGE) and "#REF!" in tok.value.upper()


#: Functions that take a LAMBDA, to which a function can also be passed by name (``BYROW(A1:C3, SUM)``).
_LAMBDA_TAKERS = {"BYROW", "BYCOL", "MAP", "REDUCE", "SCAN", "MAKEARRAY", "GROUPBY", "PIVOTBY"}


def _function_name(tok: Token) -> str:
    """``"_xlfn.LET("`` -> ``"LET"``; ``"A2:INDEX("`` -> ``"INDEX"``."""
    return tok.value[:-1].rsplit(":", 1)[-1].lstrip("@").rsplit(".", 1)[-1].upper()


def _local_names(tokens: list[Token]) -> set[str]:
    """Bare names a formula gives a meaning to itself, so they need no defined name.

    These are LET names, LAMBDA parameters, and a function passed by name as a
    whole argument of BYROW, MAP and the like. They are collected generously
    (every bare name in a LET name position or in any LAMBDA argument), because
    missing a real undefined name only makes the verdict inconclusive, while
    taking a parameter for an undefined name would fail correct work.
    """
    local: set[str] = set()
    stack: list[list] = []  # [function name, argument index]
    for i, tok in enumerate(tokens):
        if tok.subtype == Token.OPEN:
            stack.append([_function_name(tok) if tok.type == Token.FUNC else "", 0])
        elif tok.subtype == Token.CLOSE:
            if stack:
                stack.pop()
        elif tok.type == Token.SEP and tok.subtype == Token.ARG and stack:
            stack[-1][1] += 1
        elif tok.type == Token.OPERAND and tok.subtype == Token.RANGE and stack:
            func, arg = stack[-1]
            if (func == "LET" and arg % 2 == 0) or func == "LAMBDA":
                local.add(tok.value.casefold())
            elif func in _LAMBDA_TAKERS and _whole_argument(tokens, i):
                local.add(tok.value.casefold())
    return local


def _whole_argument(tokens: list[Token], i: int) -> bool:
    def boundary(tok: Token | None, edge: str) -> bool:
        return tok is not None and (tok.subtype == edge or (tok.type == Token.SEP and tok.subtype == Token.ARG))

    before = tokens[i - 1] if i > 0 else None
    after = tokens[i + 1] if i + 1 < len(tokens) else None
    return boundary(before, Token.OPEN) and boundary(after, Token.CLOSE)


def _bounds(ref: str) -> tuple[int | None, int | None, int | None, int | None]:
    min_col, min_row, max_col, max_row = range_boundaries(ref.replace("$", ""))
    if min_col is not None and max_col is not None and min_col > max_col:
        min_col, max_col = max_col, min_col
    if min_row is not None and max_row is not None and min_row > max_row:
        min_row, max_row = max_row, min_row
    return min_col, min_row, max_col, max_row


class Context:
    """Everything a rule may look at."""

    def __init__(self, case: "Case", workdir: Path, input_dir: Path, meta: dict, seed: int, engine=None,
                 cached: bool = False):
        self.case = case
        # cached=True emulates a naive grader that trusts the values stored in the file
        # instead of recalculating. Only used for ablations; never for real verdicts.
        self.cached = cached
        self.workdir = Path(workdir)
        self.input_dir = Path(input_dir)
        self.meta = meta
        self.seed = seed
        self._engine = engine
        self._wb: dict[tuple[str, bool], Any] = {}
        self._input: dict[tuple[str, bool], Any] = {}
        self._values: dict[str, dict[str | None, Values] | BaseException] = {}
        self._diagnosis: dict[tuple, str | None] = {}
        self._tmp = Path(tempfile.mkdtemp(prefix="failbook-grade-"))

    # -- files --------------------------------------------------------------------
    @property
    def primary(self) -> str:
        return self.case.outputs[0]

    @property
    def engine(self):
        if self._engine is None:
            self._engine = default_engine()
        return self._engine

    def path(self, name: str | None = None) -> Path:
        name = name or self.primary
        p = self.workdir / name
        if not p.is_file():
            raise Fail(f"expected output file {name!r} was not found")
        return p

    def wb(self, name: str | None = None, *, data_only: bool = False):
        """The agent's output workbook as saved (formulas, styles; no recalculation)."""
        name = name or self.primary
        key = (name, data_only)
        if key not in self._wb:
            try:
                self._wb[key] = openpyxl.load_workbook(self.path(name), data_only=data_only)
            except Fail:
                raise
            except Exception as exc:
                raise Fail(f"{name!r} cannot be opened as an .xlsx workbook: {exc}") from exc
        return self._wb[key]

    def ws(self, sheet: str, name: str | None = None):
        wb = self.wb(name)
        if sheet not in wb.sheetnames:
            raise Fail(f"sheet {sheet!r} is missing from {name or self.primary!r} (sheets: {', '.join(wb.sheetnames)})")
        return wb[sheet]

    def input_wb(self, name: str | None = None, *, data_only: bool = True):
        """The pristine input workbook. ``data_only=True`` gives the cached (correct) values."""
        name = name or self.case.inputs[0]
        key = (name, data_only)
        if key not in self._input:
            self._input[key] = openpyxl.load_workbook(self.input_dir / name, data_only=data_only)
        return self._input[key]

    # -- recalculated values ------------------------------------------------------
    def values(self, name: str | None = None, scenario: str | None = None) -> Values:
        """Values of the output after a full recalculation.

        ``scenario`` names a perturbation declared by the case's ``scenarios()``
        hook: the same workbook with some input cells changed. Comparing a
        scenario with the base values shows whether results are live formulas.
        """
        name = name or self.primary
        if self.cached:
            if scenario is not None:
                raise Inconclusive("the cached-value grader cannot recalculate perturbed inputs")
            return Values(self.path(name), "cached", name)
        if name not in self._values:
            try:
                self._values[name] = self._compute(name)
            except BaseException as exc:  # remember the failure for every rule
                self._values[name] = exc
        got = self._values[name]
        if isinstance(got, BaseException):
            raise got
        if scenario not in got:
            for key in (("error", scenario), ("error", "__all__")):
                err = got.get(key)
                if isinstance(err, BaseException):
                    raise err
            raise Inconclusive(f"scenario {scenario!r} was not prepared for {name!r}")
        return got[scenario]

    def _compute(self, name: str) -> dict:
        src = self.path(name)
        self.wb(name)  # fail early with a clear message if the file is not a workbook
        work = self._tmp / Path(name).stem
        if work.exists():
            shutil.rmtree(work)
        (work / "in").mkdir(parents=True)
        files: dict[str | None, Path] = {}
        base = work / "in" / "__base__.xlsx"
        shutil.copyfile(src, base)
        files[None] = base
        errors: dict[Any, BaseException] = {}
        for label, overrides in self._scenarios_for(name, errors).items():
            wb = openpyxl.load_workbook(src)
            for ref, value in overrides.items():
                sheet, coord = X.split_ref(ref)
                if sheet not in wb.sheetnames:
                    errors[("error", label)] = Fail(f"cannot perturb {ref}: sheet {sheet!r} not found")
                    break
                wb[sheet][coord].value = value
            else:
                path = work / "in" / f"{label}.xlsx"
                wb.save(path)
                files[label] = path
        mapping = self.engine.recalc(files.values(), work / "out")
        result: dict[Any, Any] = {label: Values(mapping[p.resolve()], label or "base", name) for label, p in files.items()}
        result.update(errors)
        return result

    def _scenarios_for(self, name: str, errors: dict) -> dict[str, dict[str, Any]]:
        hook = getattr(self.case.module, "scenarios", None)
        if hook is None:
            return {}
        try:
            declared = hook(self) or {}
        except (Fail, Inconclusive, AssertionError) as exc:
            # Rules that need a scenario will report this; base values still work.
            err = exc if not isinstance(exc, AssertionError) else Fail(str(exc) or "could not prepare scenarios")
            errors[("error", "__all__")] = err
            return {}
        out = {}
        for label, spec in declared.items():
            if isinstance(spec, tuple):
                file, overrides = spec
            else:
                file, overrides = self.primary, spec
            if file == name:
                out[label] = overrides
        return out

    # -- typed accessors with honest error reporting --------------------------------
    def value(self, values: Values, ref: str) -> Any:
        v = values[ref]
        if X.is_error(v):
            self._raise_for_error(values, ref, v)
        return v

    def num(self, values: Values, ref: str, what: str | None = None) -> float:
        """A numeric result, or a precise Fail/Inconclusive explaining why not."""
        v = values[ref]
        label = f"{what} ({ref})" if what else ref
        if X.is_error(v):
            self._raise_for_error(values, ref, v, label)
        if v is None:
            raise Fail(f"{label} is empty")
        if isinstance(v, bool) or not X.is_number(v):
            raise Fail(f"{label} should be a number but is {X.show(v)}")
        return float(v)

    def _raise_for_error(self, values: Values, ref: str, v: str, label: str | None = None):
        label = label or ref
        if v.strip().upper() == "#NAME?":
            cause = self.explain_name_error(values, ref)
            if cause:
                raise Fail(f"{label} evaluates to #NAME? because {cause}")
            raise Inconclusive(
                f"{label} evaluates to #NAME? and no broken reference was found in its formula or in the "
                "formulas it depends on; it may use a function the recalculation engine does not support (or "
                "one saved without its _xlfn. prefix). Re-check with a newer LibreOffice or Excel before "
                "counting this as a failure."
            )
        raise Fail(f"{label} evaluates to {v}")

    def explain_name_error(self, values: Values, ref: str) -> str | None:
        """Name a broken reference that explains the ``#NAME?`` in ``ref``, or return None.

        Only the cell's own formula and the cells it depends on that also evaluate to
        ``#NAME?`` are examined, so an unrelated problem elsewhere in the workbook is
        never blamed. A cause is reported only when it is a fact about the file: a
        reference to a missing sheet or table, a literal ``#REF!``, a defined name that
        points to a missing sheet or contains ``#REF!``, an undefined name, or an error
        value stored as a constant. LET/LAMBDA parameters, structured references to
        existing tables and function names are not names that need defining.
        """
        key = (values.source, getattr(values, "label", None), ref)
        if key not in self._diagnosis:
            self._diagnosis[key] = self._explain_name_error(values, ref)
        return self._diagnosis[key]

    def _explain_name_error(self, values: Values, ref: str) -> str | None:
        wb = self.wb(values.source)
        refs = _References(wb)
        calc = getattr(values, "wb", None)  # recalculated values, to follow the error to its origin
        sheet, coord = X.split_ref(ref)
        start = (sheet, *X.coord_to_rc(coord))
        queue, seen, budget = deque([start]), {start}, _PRECEDENT_BUDGET
        while queue:
            title, row, col = queue.popleft()
            if title not in wb.sheetnames:
                continue
            cell = wb[title]._cells.get((row, col))
            raw = cell.value if cell is not None else None
            where = X.make_ref(title, (row, col))
            formula = X.formula_text(raw)
            if formula is None:
                if X.is_error(raw):
                    return f"{where} holds the error value {raw} as a constant, not a formula"
                continue
            problem = refs.problem(title, formula)
            if problem:
                return f"{where} {problem}: {formula}"
            if calc is None:
                continue
            for p_sheet, min_col, min_row, max_col, max_row in refs.precedents(title, formula):
                if p_sheet not in calc.sheetnames:
                    continue
                cws = calc[p_sheet]
                max_row = min(max_row or cws.max_row, cws.max_row)
                max_col = min(max_col or cws.max_column, cws.max_column)
                for r in range(min_row or 1, max_row + 1):
                    for c in range(min_col or 1, max_col + 1):
                        budget -= 1
                        if budget < 0:
                            return None
                        got = cws._cells.get((r, c))
                        node = (p_sheet, r, c)
                        if got is not None and _is_name_error(got.value) and node not in seen:
                            seen.add(node)
                            queue.append(node)
        return None

    def cleanup(self) -> None:
        shutil.rmtree(self._tmp, ignore_errors=True)


def run_rules(case: "Case", ctx: Context, kinds: Iterable[str] | None = None) -> list[RuleResult]:
    """Run the case's rules (optionally only those of the given kinds)."""
    wanted = set(kinds) if kinds is not None else None
    results = []
    for spec in case.rules:
        if wanted is not None and spec.kind not in wanted:
            continue
        try:
            spec.func(ctx)
            results.append(RuleResult(spec.id, spec.title, PASS, kind=spec.kind))
        except Fail as exc:
            results.append(RuleResult(spec.id, spec.title, FAIL, str(exc), kind=spec.kind))
        except AssertionError as exc:
            results.append(RuleResult(spec.id, spec.title, FAIL, str(exc) or "assertion failed", kind=spec.kind))
        except Inconclusive as exc:
            results.append(RuleResult(spec.id, spec.title, INCONCLUSIVE, str(exc), kind=spec.kind))
        except Exception as exc:  # a bug in the checker must never fail the agent
            tb = traceback.format_exception_only(type(exc), exc)[-1].strip()
            results.append(RuleResult(spec.id, spec.title, INCONCLUSIVE, f"checker error: {tb}", kind=spec.kind))
    return results


def verdict_of(results: Iterable[RuleResult]) -> str:
    results = list(results)
    if any(r.status == FAIL for r in results):
        return FAIL
    if any(r.status == INCONCLUSIVE for r in results):
        return INCONCLUSIVE
    return PASS
