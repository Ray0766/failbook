"""Every case is complete and internally consistent (no LibreOffice needed)."""

import json

import pytest

from failbook.case import all_cases
from failbook.checks import RULE_KINDS

CASES = all_cases()


def test_there_are_cases():
    assert 10 <= len(CASES) <= 20


@pytest.mark.parametrize("case", CASES, ids=lambda c: c.id)
def test_case_is_complete(case):
    spec = case.spec
    for key in ("expected", "observed", "explanation", "claim", "failing_rules"):
        assert spec["failure"].get(key), f"failure.{key} missing"
    for key in ("expected_zh", "observed_zh", "explanation_zh", "claim_zh"):
        assert spec["failure"].get(key), f"failure.{key} missing"
    assert spec["classification"]["category"] in {"model", "tool", "environment", "evaluator", "undetermined"}
    assert spec["classification"]["evidence"] in {"reconstructed", "observed"}
    assert spec["limitations"] and all(item.get("en") and item.get("zh") for item in spec["limitations"])
    assert spec["inputs"]["source"] and spec["inputs"]["license"]
    ids = {r.id for r in case.rules}
    assert set(case.expected_failures) <= ids, "failing_rules must name existing rules"
    for r in case.rules:
        assert r.kind in RULE_KINDS
        assert r.requirement.strip(), f"{r.id} must quote the task sentence it checks"
        assert r.title_zh, f"{r.id} needs a Chinese title"
    for name in case.inputs:
        assert (case.input_dir / name).is_file()
    for name in case.outputs:
        assert (case.failure_dir / name).is_file()
    grade = json.loads((case.failure_dir / "grade.json").read_text(encoding="utf-8"))
    failed = {r["id"] for r in grade["rules"] if r["status"] == "fail"}
    assert set(case.expected_failures) <= failed, "committed grading must fail the listed rules"


@pytest.mark.parametrize("case", CASES, ids=lambda c: c.id)
def test_tasks_render_in_both_languages(case):
    meta = case.params(0)
    for lang in ("en", "zh"):
        text = case.prompt(meta, lang)
        assert "{{" not in text and len(text) > 100


@pytest.mark.parametrize("case", CASES, ids=lambda c: c.id)
def test_variants_are_deterministic_and_different(case):
    assert case.params(1) == case.params(1)
    assert case.params(0) != case.params(1)
