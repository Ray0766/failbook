"""Grading an output directory, and verifying that a case's checker is trustworthy."""

from __future__ import annotations

import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from . import __version__
from .case import Case
from .checks import FAIL, PASS, Context, GradeReport, run_rules, verdict_of
from .engine import default_engine


def grade(case: Case, workdir: Path, seed: int = 0, *, engine=None, input_dir: Path | None = None,
          kinds=None, cached: bool = False) -> GradeReport:
    """Grade the files an agent left in ``workdir`` for the given case variant.

    ``kinds`` and ``cached`` exist for ablations (weaker graders used as baselines);
    real verdicts always use every rule and full recalculation.
    """
    engine = engine or default_engine()
    tmp = None
    try:
        if input_dir is None:
            tmp = Path(tempfile.mkdtemp(prefix="failbook-input-"))
            meta = case.inputs_for(seed, tmp, engine=engine)
            input_dir = tmp
        else:
            meta = case.params(seed)
        ctx = Context(case, Path(workdir), Path(input_dir), meta, seed, engine=engine, cached=cached)
        try:
            results = run_rules(case, ctx, kinds)
        finally:
            ctx.cleanup()
        return GradeReport(
            case=case.id,
            case_version=case.version,
            seed=seed,
            verdict=verdict_of(results),
            rules=results,
            engine=engine.describe(),
            failbook_version=__version__,
        )
    finally:
        if tmp is not None:
            shutil.rmtree(tmp, ignore_errors=True)


def prepare(case: Case, workdir: Path, seed: int = 0, *, lang: str = "en", engine=None,
            write_task: bool = True) -> dict[str, Any]:
    """Create a fresh working directory with the inputs (and TASK.md) for one attempt."""
    workdir = Path(workdir)
    workdir.mkdir(parents=True, exist_ok=True)
    meta = case.inputs_for(seed, workdir, engine=engine)
    if write_task:
        (workdir / "TASK.md").write_text(case.prompt(meta, lang), encoding="utf-8")
    return meta


@dataclass
class VerifyOutcome:
    case: str
    seed: int
    reference: GradeReport
    failure: GradeReport
    problems: list[str] = field(default_factory=list)
    #: ``alt_*.py``: other valid solutions, graded as written.
    alternatives: dict[str, GradeReport] = field(default_factory=dict)
    #: ``fail_*.py``: other wrong solutions (regression fixtures), each expected to be graded fail.
    fixtures: dict[str, GradeReport] = field(default_factory=dict)
    #: The reference and every alternative after LibreOffice opened and re-saved their output.
    resaved: dict[str, GradeReport] = field(default_factory=dict)

    @property
    def ok(self) -> bool:
        return not self.problems


def _resave(workdirs: dict[str, Path], outputs: list[str], tmp: Path, engine) -> dict[str, Path]:
    """Open and re-save each solution's output files with LibreOffice, in one batch.

    ``workdirs`` maps a solution kind to its working directory; the result maps the
    kind to a directory holding the re-saved files under their original names.
    """
    staging, out = tmp / "_resave_in", tmp / "_resave_out"
    staging.mkdir(parents=True)
    origin: dict[Path, tuple[str, str]] = {}
    for kind, work in workdirs.items():
        for name in outputs:
            src = work / name
            if src.is_file():
                staged = staging / f"{kind}__{name}"
                shutil.copyfile(src, staged)
                origin[staged.resolve()] = (kind, name)
    mapping = engine.recalc(list(origin), out) if origin else {}
    dirs = {}
    for kind in workdirs:
        dirs[kind] = tmp / "resaved" / kind
        dirs[kind].mkdir(parents=True)
    for staged, saved in mapping.items():
        kind, name = origin[Path(staged).resolve()]
        shutil.copyfile(saved, dirs[kind] / name)
    return dirs


def verify(case: Case, seeds: list[int], *, engine=None) -> list[VerifyOutcome]:
    """Prove the checker is fair and sharp for each seed.

    * the reference solution must pass every rule (no false failures),
    * every alternative solution (``alt_*.py``: other valid ways to do the task)
      must pass too,
    * the output of the reference and of every alternative, opened and re-saved by
      a spreadsheet application, must still pass (the checks do not depend on how a
      file was written),
    * the untouched input (an agent that did nothing) must fail,
    * the failure reproduction must fail, and fail on the rules the case says it
      fails (so the case catches the failure for the right reason), and
    * every failure fixture (``fail_*.py``: other wrong solutions) must be graded
      fail; an inconclusive verdict is a problem too, because it would not be
      counted against an agent that made the same mistake.
    """
    engine = engine or default_engine()
    outcomes = []
    for seed in seeds:
        with tempfile.TemporaryDirectory(prefix=f"failbook-verify-{case.id}-") as tmp:
            tmp = Path(tmp)
            inputs = tmp / "input"
            meta = case.build(seed, inputs, engine=engine)
            alternatives, fixtures = case.alternatives(), case.failure_fixtures()
            reports: dict[str, GradeReport] = {}
            for kind in ("reference", "failure", *alternatives, *fixtures):
                work = tmp / "solved" / kind
                shutil.copytree(inputs, work)
                case.solve(kind, work, meta)
                reports[kind] = grade(case, work, seed, engine=engine, input_dir=inputs)
            problems = []
            ref, fail = reports["reference"], reports["failure"]
            for kind in ("reference", *alternatives):
                for r in reports[kind].rules:
                    if r.status != PASS:
                        problems.append(f"{kind}: rule {r.id} is {r.status}: {r.detail}")
            if fail.verdict != FAIL:
                problems.append(f"failure reproduction was graded {fail.verdict}, expected fail")
            failed_ids = {r.id for r in fail.rules if r.status == FAIL}
            for expected in case.expected_failures:
                if expected not in failed_ids:
                    problems.append(f"failure reproduction should fail rule {expected} but it did not")
            if not case.expected_failures:
                problems.append("case.yaml must list failure.failing_rules")
            for kind in fixtures:
                rep = reports[kind]
                if rep.verdict != FAIL:
                    undecided = "; ".join(f"{r.id} is {r.status}: {r.detail}" for r in rep.rules
                                          if r.status != PASS)
                    problems.append(f"{kind} (a wrong solution) was graded {rep.verdict}, expected fail"
                                    + (f" ({undecided})" if undecided else ""))
            # Doing nothing must never pass: grade the untouched inputs.
            noop = grade(case, inputs, seed, engine=engine, input_dir=inputs)
            if noop.verdict != FAIL:
                problems.append(f"untouched input was graded {noop.verdict}; an empty attempt must fail")
            # Re-saved by LibreOffice, like a user opening and saving the file.
            resaved_dirs = _resave({k: tmp / "solved" / k for k in ("reference", *alternatives)},
                                   case.outputs, tmp, engine)
            resaved = {}
            for kind, work in resaved_dirs.items():
                rs = grade(case, work, seed, engine=engine, input_dir=inputs)
                resaved[kind] = rs
                for r in rs.rules:
                    if r.status != PASS:
                        problems.append(f"{kind} re-saved by LibreOffice: rule {r.id} is {r.status}: {r.detail}")
            outcomes.append(VerifyOutcome(
                case.id, seed, ref, fail, problems,
                alternatives={k: reports[k] for k in alternatives},
                fixtures={k: reports[k] for k in fixtures},
                resaved=resaved,
            ))
    return outcomes
