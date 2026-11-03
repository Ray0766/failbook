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
    alternatives: dict[str, GradeReport] = field(default_factory=dict)

    @property
    def ok(self) -> bool:
        return not self.problems


def verify(case: Case, seeds: list[int], *, engine=None) -> list[VerifyOutcome]:
    """Prove the checker is fair and sharp for each seed.

    * the reference solution must pass every rule (no false failures),
    * the untouched input (an agent that did nothing) must fail,
    * the failure reproduction must fail, and fail on the rules the case says
      it fails (so the case catches the failure for the right reason),
    * the reference output, opened and re-saved by a spreadsheet application,
      must still pass (the checks do not depend on how a file was written), and
    * every alternative solution (``alt_*.py``: other valid ways to do the task)
      must pass too.
    """
    engine = engine or default_engine()
    outcomes = []
    for seed in seeds:
        with tempfile.TemporaryDirectory(prefix=f"failbook-verify-{case.id}-") as tmp:
            tmp = Path(tmp)
            inputs = tmp / "input"
            meta = case.build(seed, inputs, engine=engine)
            reports = {}
            for kind in ("reference", "failure"):
                work = tmp / kind
                shutil.copytree(inputs, work)
                case.solve(kind, work, meta)
                reports[kind] = grade(case, work, seed, engine=engine, input_dir=inputs)
            problems = []
            ref, fail = reports["reference"], reports["failure"]
            for r in ref.rules:
                if r.status != PASS:
                    problems.append(f"reference: rule {r.id} is {r.status}: {r.detail}")
            if fail.verdict != FAIL:
                problems.append(f"failure reproduction was graded {fail.verdict}, expected fail")
            failed_ids = {r.id for r in fail.rules if r.status == FAIL}
            for expected in case.expected_failures:
                if expected not in failed_ids:
                    problems.append(f"failure reproduction should fail rule {expected} but it did not")
            if not case.expected_failures:
                problems.append("case.yaml must list failure.failing_rules")
            # Doing nothing must never pass: grade the untouched inputs.
            noop = grade(case, inputs, seed, engine=engine, input_dir=inputs)
            if noop.verdict != FAIL:
                problems.append(f"untouched input was graded {noop.verdict}; an empty attempt must fail")
            # Re-saved by LibreOffice, like a user opening and saving the file.
            resaved = tmp / "resaved"
            names = [tmp / "reference" / n for n in case.outputs]
            mapping = engine.recalc(names, resaved / "_out")
            (resaved / "work").mkdir(parents=True)
            for src, out in mapping.items():
                shutil.copyfile(out, resaved / "work" / src.name)
            rs = grade(case, resaved / "work", seed, engine=engine, input_dir=inputs)
            for r in rs.rules:
                if r.status != PASS:
                    problems.append(f"reference re-saved by LibreOffice: rule {r.id} is {r.status}: {r.detail}")
            alternatives = {}
            for alt in case.alternatives():
                work = tmp / alt
                shutil.copytree(inputs, work)
                case.solve(alt, work, meta)
                rep = grade(case, work, seed, engine=engine, input_dir=inputs)
                alternatives[alt] = rep
                for r in rep.rules:
                    if r.status != PASS:
                        problems.append(f"{alt}: rule {r.id} is {r.status}: {r.detail}")
            outcomes.append(VerifyOutcome(case.id, seed, ref, fail, problems, alternatives))
    return outcomes
