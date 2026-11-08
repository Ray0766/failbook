"""`failbook verify` and the solutions it runs: failure fixtures (fail_*.py) and the re-save check.

A toy case and a stand-in engine (it "recalculates" by copying) keep these fast and independent of
LibreOffice; the toy rule reads the saved cell directly, so it needs no recalculation.
"""

import shutil
import textwrap
from pathlib import Path

import pytest

from failbook.case import CaseError, load_case_dir
from failbook.docs import case_readme
from failbook.grade import verify

CASE_PY = '''
import openpyxl
from failbook.checks import Inconclusive, require, rule

def params(seed):
    return {"n": 3 + seed}

def build(meta, outdir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for i in range(meta["n"]):
        ws.cell(i + 1, 1, i + 1)
    wb.save(outdir / "book.xlsx")

@rule("total-correct", "B1 holds the total of column A", requirement="Put the total of column A in B1.",
      title_zh="B1 为 A 列合计")
def total_correct(ctx):
    got = ctx.ws("S")["B1"].value
    if got == "?":
        raise Inconclusive("cannot tell")
    want = sum(range(1, ctx.meta["n"] + 1))
    require(got == want, f"B1 is {got!r}, expected {want}")
'''

SOLUTION = '''
import openpyxl

def solve(workdir, meta):
    path = workdir / "book.xlsx"
    wb = openpyxl.load_workbook(path)
    total = sum(range(1, meta["n"] + 1))
    wb["S"]["B1"] = {value}
    wb.save(path)
'''


def make_case(root: Path, solutions: dict[str, str]):
    d = root / "xlsx-901-toy"
    d.mkdir()
    (d / "case.yaml").write_text(textwrap.dedent("""\
        id: xlsx-901
        slug: toy
        title: Toy case
        files: {inputs: [book.xlsx], outputs: [book.xlsx]}
        failure: {failing_rules: [total-correct]}
        """), encoding="utf-8")
    (d / "task.md").write_text("Put the total of column A in B1.\n", encoding="utf-8")
    (d / "case.py").write_text(CASE_PY, encoding="utf-8")
    for kind, value in {"reference": "total", "failure": "0", **solutions}.items():
        (d / f"{kind}.py").write_text(SOLUTION.format(value=value), encoding="utf-8")
    return load_case_dir(d)


class CopyEngine:
    """Stands in for LibreOffice. ``break_on_resave`` lists solution kinds whose re-saved copy loses B1."""

    def __init__(self, break_on_resave=()):
        self.break_on_resave = tuple(break_on_resave)
        self.batches: list[list[str]] = []

    def recalc(self, files, outdir):
        import openpyxl

        outdir = Path(outdir)
        outdir.mkdir(parents=True, exist_ok=True)
        files = [Path(f).resolve() for f in files]
        self.batches.append([f.name for f in files])
        out = {}
        for f in files:
            target = outdir / (f.stem + ".xlsx")
            shutil.copyfile(f, target)
            if f.name.startswith(tuple(f"{k}__" for k in self.break_on_resave)):
                wb = openpyxl.load_workbook(target)
                wb["S"]["B1"] = None
                wb.save(target)
            out[f] = target
        return out

    def describe(self):
        return "copy (test stand-in)"


def test_fixtures_are_listed_and_loadable(tmp_path):
    case = make_case(tmp_path, {"fail_off_by_one": "total + 1", "fail_blank": "None", "alt_same": "total"})
    assert case.failure_fixtures() == ["fail_blank", "fail_off_by_one"]
    assert case.alternatives() == ["alt_same"]
    assert callable(case.solution("fail_off_by_one").solve)
    for bad in ("fail_", "fail_../reference", "other", "alt_x/y"):
        with pytest.raises(CaseError, match="unknown solution kind"):
            case.solution(bad)
    with pytest.raises(CaseError, match="fail_missing.py is missing"):
        case.solution("fail_missing")


def test_verify_requires_every_fixture_to_fail(tmp_path):
    case = make_case(tmp_path, {"fail_off_by_one": "total + 1", "fail_correct_after_all": "total",
                                "fail_unsure": "'?'"})
    [outcome] = verify(case, [0], engine=CopyEngine())
    assert set(outcome.fixtures) == {"fail_off_by_one", "fail_correct_after_all", "fail_unsure"}
    assert outcome.fixtures["fail_off_by_one"].verdict == "fail"
    assert not outcome.ok
    problems = "\n".join(outcome.problems)
    assert "fail_off_by_one" not in problems
    assert "fail_correct_after_all (a wrong solution) was graded pass, expected fail" in problems
    # Inconclusive would not count against an agent that made the same mistake, so it is a problem too.
    assert "fail_unsure (a wrong solution) was graded inconclusive, expected fail" in problems


def test_verify_resaves_the_reference_and_every_alternative(tmp_path):
    case = make_case(tmp_path, {"alt_same": "total", "alt_fragile": "total", "fail_off_by_one": "total + 1"})
    engine = CopyEngine(break_on_resave=["alt_fragile"])
    [outcome] = verify(case, [0], engine=engine)
    assert set(outcome.resaved) == {"reference", "alt_same", "alt_fragile"}
    assert outcome.alternatives["alt_fragile"].verdict == "pass"  # passes as written ...
    assert outcome.resaved["alt_fragile"].verdict == "fail"       # ... but not after a re-save
    assert outcome.problems == [
        "alt_fragile re-saved by LibreOffice: rule total-correct is fail: B1 is None, expected 6"
    ]
    # Re-saving happens in one batch, and failure fixtures are not re-saved.
    resave = [b for b in engine.batches if any("__" in n for n in b)]
    assert len(resave) == 1
    assert sorted(resave[0]) == ["alt_fragile__book.xlsx", "alt_same__book.xlsx", "reference__book.xlsx"]


def test_verify_ok_case(tmp_path):
    case = make_case(tmp_path, {"alt_same": "total", "fail_off_by_one": "total + 1"})
    outcomes = verify(case, [0, 1], engine=CopyEngine())
    assert all(o.ok for o in outcomes), [o.problems for o in outcomes]


def test_case_readme_names_the_fixtures(tmp_path):
    case = make_case(tmp_path, {"alt_same": "total", "fail_off_by_one": "total + 1"})
    text = case_readme(case)
    assert ("[`reference.py`](reference.py) and [`alt_same.py`](alt_same.py) must pass, also after LibreOffice "
            "opens and re-saves their output") in text
    assert "[`fail_off_by_one.py`](fail_off_by_one.py) (another wrong solution) must be graded fail" in text

    (case.dir / "fail_off_by_one.py").unlink()
    (case.dir / "alt_same.py").unlink()
    text = case_readme(case)
    assert "[`reference.py`](reference.py) must pass, also after LibreOffice opens and re-saves its output" in text
    assert "wrong solution" not in text
