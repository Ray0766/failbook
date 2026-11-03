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
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Any, Callable, Iterable

import openpyxl
from openpyxl.formula.tokenizer import Token, Tokenizer

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
    "preserve": "nothing the task did not ask for was changed",
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
    actually asked to do; see docs/fairness.md.
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
        self._diagnosis: dict[str, str | None] = {}
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
        source = values.source
        if v.strip().upper() == "#NAME?":
            broken = self.broken_references(source)
            if broken:
                raise Fail(f"{label} evaluates to #NAME? because {broken}")
            raise Inconclusive(
                f"{label} evaluates to #NAME? and no broken reference was found; the formula may use a "
                "function the recalculation engine does not support (or that was saved without its "
                "_xlfn. prefix). Re-check with a newer LibreOffice or Excel before counting this as a failure."
            )
        raise Fail(f"{label} evaluates to {v}")

    def broken_references(self, name: str | None = None) -> str | None:
        """Describe the first formula that points to a missing sheet or name, if any."""
        name = name or self.primary
        if name in self._diagnosis:
            return self._diagnosis[name]
        wb = self.wb(name)
        sheets = set(wb.sheetnames)
        names = set()
        try:
            names |= set(wb.defined_names.keys())
        except Exception:  # pragma: no cover - older openpyxl
            pass
        for ws in wb.worksheets:
            try:
                names |= set(ws.defined_names.keys())
            except Exception:  # pragma: no cover
                pass
        lowered_sheets = {s.casefold() for s in sheets}
        lowered_names = {n.casefold() for n in names}
        found = None
        for ws in wb.worksheets:
            for row, col, value in X.used_cells(ws):
                formula = X.formula_text(value)
                if not formula:
                    continue
                where = X.make_ref(ws.title, (row, col))
                if "#REF!" in formula.upper():
                    found = f"{where} contains a broken reference: {formula}"
                    break
                try:
                    tokens = Tokenizer(formula).items
                except Exception:
                    continue
                for tok in tokens:
                    if tok.type != Token.OPERAND or tok.subtype != Token.RANGE:
                        continue
                    text = tok.value
                    if "!" in text:
                        sheet = text.rsplit("!", 1)[0]
                        if sheet.startswith("'") and sheet.endswith("'"):
                            sheet = sheet[1:-1].replace("''", "'")
                        if sheet.startswith("["):
                            continue  # external workbook reference
                        if sheet.casefold() not in lowered_sheets:
                            found = f"{where} refers to a sheet that does not exist ({sheet!r}): {formula}"
                            break
                    elif not _CELLISH.match(text) and text.casefold() not in lowered_names:
                        found = f"{where} refers to an undefined name {text!r}: {formula}"
                        break
                if found:
                    break
            if found:
                break
        self._diagnosis[name] = found
        return found

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
