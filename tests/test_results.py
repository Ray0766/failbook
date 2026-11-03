import json

from failbook.checks import GradeReport, RuleResult
from failbook.harbor import write_reward
from failbook.report import render_markdown, validate_result
from failbook.runner import auto_classification, run_status


def good_result():
    return {"schema": "failbook.result/1", "case": "xlsx-001", "case_version": 1, "seed": 0,
            "agent": {"name": "demo"}, "model": "m", "harness": "manual", "started_at": "2026-11-03T00:00:00+00:00",
            "verdict": "fail", "rules": [{"id": "a", "status": "fail"}]}


def test_result_validation():
    assert validate_result(good_result()) == []
    bad = good_result()
    del bad["verdict"]
    assert validate_result(bad)
    bad = good_result()
    bad["rules"] = [{"id": "a", "status": "maybe"}]
    assert validate_result(bad)


def test_run_status_and_classification_keep_environment_errors_apart():
    assert run_status(0, False) == "completed"
    assert run_status(None, True) == "timeout"
    assert run_status(2, False) == "agent-error"
    assert auto_classification("fail", "timeout") == "agent-or-environment-error"
    assert auto_classification("fail", "completed") == "task-failure"
    assert auto_classification("inconclusive", "completed") == "evaluator-inconclusive"
    assert auto_classification("pass", "completed") == "none"


def test_report_has_no_ranking():
    text = render_markdown([dict(good_result(), case_version=1)])
    assert "All cases" not in text and "no overall score" in text


def test_harbor_reward_is_flat_and_numeric(tmp_path):
    report = GradeReport("xlsx-001", 1, 0, "inconclusive", [RuleResult("a", "a", "inconclusive")], "lo", "0")
    write_reward(report, tmp_path)
    data = json.loads((tmp_path / "reward.json").read_text())
    assert all(isinstance(v, (int, float)) for v in data.values())
    assert data["reward"] == 0 and data["inconclusive"] == 1
    assert (tmp_path / "reward.txt").read_text().strip() == "0"
