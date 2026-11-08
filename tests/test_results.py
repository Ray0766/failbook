import json
import os
import shutil
import time
from pathlib import Path

import pytest

from failbook import engine as engine_mod
from failbook.case import load_case, repo_root
from failbook.checks import GradeReport, RuleResult
from failbook.cli import main
from failbook.harbor import write_grader_error, write_reward
from failbook.report import render_markdown, validate_result
from failbook.runner import PRESETS, auto_classification, build_argv, run_agent, run_status, sha256_file

from conftest import needs_libreoffice

POSIX = pytest.mark.skipif(not hasattr(os, "killpg"), reason="process groups are POSIX only")


def good_result(**changes):
    r = {"schema": "failbook.result/1", "case": "xlsx-001", "case_version": load_case("xlsx-001").version, "seed": 0,
         "agent": {"name": "demo"}, "model": "m", "harness": "manual", "started_at": "2026-11-03T00:00:00+00:00",
         "verdict": "fail", "rules": [{"id": "a", "status": "fail"}]}
    r.update(changes)
    return r


def test_result_validation():
    assert validate_result(good_result()) == []
    bad = good_result()
    del bad["verdict"]
    assert validate_result(bad)
    bad = good_result()
    bad["rules"] = [{"id": "a", "status": "maybe"}]
    assert validate_result(bad)
    assert validate_result(good_result(run_status="crashed"))


def test_run_status_and_classification_keep_environment_errors_apart():
    assert run_status(0, False) == "completed"
    assert run_status(None, True) == "timeout"
    assert run_status(2, False) == "agent-error"
    assert auto_classification("fail", "timeout") == "agent-or-environment-error"
    assert auto_classification("fail", "completed") == "task-failure"
    assert auto_classification("inconclusive", "completed") == "evaluator-inconclusive"
    assert auto_classification("pass", "completed") == "none"


def test_report_has_no_ranking():
    text = render_markdown([good_result()])
    assert "All cases" not in text and "no overall score" in text


def _row(text, case_id):
    return next(line for line in text.splitlines() if line.startswith(f"| [{case_id}]"))


def test_report_keeps_run_problems_and_inconclusive_out_of_failures():
    crashy = dict(model="crashy")
    unsure = dict(model="unsure")
    text = render_markdown([
        good_result(**crashy, run_status="agent-error", exit_code=1),
        good_result(**crashy, run_status="timeout", timed_out=True),
        good_result(**crashy, verdict="pass", run_status="completed"),
        good_result(**unsure, verdict="inconclusive"),
        good_result(**unsure, verdict="inconclusive"),
    ])
    row = _row(text, "xlsx-001")
    assert "✅ 1/1 (1 timeout, 1 agent-error)" in row  # the crashes are not failed attempts
    assert "❔ 2 inconclusive" in row and "❌" not in row


def test_report_shows_older_case_versions_separately():
    current = load_case("xlsx-001").version
    text = render_markdown([good_result(), good_result(case_version=current - 1, verdict="pass")])
    assert "❌ 0/1" in _row(text, "xlsx-001")  # only the current version counts in the main table
    older = text.split("## Results for older case versions", 1)[1]
    assert f"| {current - 1} | {current} | ✅ 1/1 |" in older


def test_report_without_results_does_not_promise_harbor_writes_results():
    text = render_markdown([])
    assert "No retest results" in text and "or a Harbor job" not in text and "grade --record" in text


def test_report_exits_nonzero_on_invalid_result_files(tmp_path, capsys):
    (tmp_path / "xlsx-001").mkdir()
    (tmp_path / "xlsx-001" / "ok.json").write_text(json.dumps(good_result()))
    assert main(["report", "--results", str(tmp_path), "--out", str(tmp_path / "R.md")]) == 0
    (tmp_path / "xlsx-001" / "bad.json").write_text('{"schema": "failbook.result/1", "verdict": "passed"}')
    assert main(["report", "--results", str(tmp_path), "--out", str(tmp_path / "R.md")]) == 1
    assert "invalid result" in capsys.readouterr().err


def test_harbor_reward_is_flat_and_numeric(tmp_path):
    report = GradeReport("xlsx-001", 1, 0, "inconclusive", [RuleResult("a", "a", "inconclusive")], "lo", "0")
    write_reward(report, tmp_path / "inconclusive")
    write_grader_error(tmp_path / "crash")
    data = json.loads((tmp_path / "inconclusive" / "reward.json").read_text())
    crash = json.loads((tmp_path / "crash" / "reward.json").read_text())
    assert all(isinstance(v, (int, float)) for v in data.values())
    assert data.keys() == crash.keys()  # Harbor averages per key and counts a missing key as 0
    assert data["reward"] == 0 and data["inconclusive"] == 1 and data["grader_error"] == 0
    assert crash["reward"] == 0 and crash["grader_error"] == 1 and crash["inconclusive"] == 0
    assert (tmp_path / "inconclusive" / "reward.txt").read_text().strip() == "0"


def _argv(agent, model="m"):
    return build_argv(PRESETS[agent], None, prompt="PROMPT", prompt_file=Path("TASK.md"), workdir=Path("w"),
                      model=model)


def test_preset_model_options_never_split_the_prompt_from_its_flag():
    gemini = _argv("gemini-cli")
    assert gemini[gemini.index("-p") + 1] == "PROMPT" and gemini[-2:] == ["--model", "m"]
    codex = _argv("codex")
    assert codex[-1] == "PROMPT" and codex[-3:-1] == ["--model", "m"]
    claude = _argv("claude-code")
    assert claude[claude.index("-p") + 1] == "PROMPT" and "--model" in claude
    assert "--model" not in _argv("gemini-cli", model="") and "{model_args}" not in _argv("codex", model="")


def test_custom_placeholders_are_absolute(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    argv = build_argv(None, "agent {prompt_file} {workdir}", prompt="x", prompt_file=Path("r/TASK.md"),
                      workdir=Path("r/work"), model="")
    assert argv[1] == str(tmp_path / "r" / "TASK.md") and argv[2] == str(tmp_path / "r" / "work")


def test_runs_started_in_the_same_second_get_their_own_directories(tmp_path, capsys):
    args = ["run", "xlsx-006", "--agent", "custom", "--cmd", "agent {prompt_file}", "--dry-run",
            "--runs-dir", str(tmp_path / "runs"), "--model", "a"]
    assert main(args) == 0 and main(args) == 0
    out = capsys.readouterr().out
    dirs = sorted((tmp_path / "runs").iterdir())
    assert len(dirs) == 2
    for d in dirs:
        assert (d / "TASK.md").is_file() and f"agent {d / 'TASK.md'}" in out


def test_runs_dir_inside_the_checkout_is_refused(capsys):
    root = repo_root()
    assert root is not None
    code = main(["run", "xlsx-006", "--agent", "custom", "--cmd", "agent", "--dry-run",
                 "--runs-dir", str(root / "runs-test-should-not-exist")])
    assert code == 2 and "reference solutions" in capsys.readouterr().err
    assert not (root / "runs-test-should-not-exist").exists()


def test_default_runs_dir_is_outside_the_checkout(capsys):
    assert main(["run", "xlsx-006", "--agent", "custom", "--cmd", "agent {workdir}", "--dry-run"]) == 0
    line = next(l for l in capsys.readouterr().out.splitlines() if l.startswith("[dry-run]"))
    workdir = Path(line.split()[-1])
    try:
        assert workdir.is_absolute() and repo_root() not in workdir.parents
    finally:
        shutil.rmtree(workdir.parents[1], ignore_errors=True)


@POSIX
def test_timeout_kills_the_agents_whole_process_group(tmp_path):
    writer = f"(for i in $(seq 1 30); do sleep 0.2; echo $i >> {tmp_path}/late.txt; done) & sleep 60"
    started = time.monotonic()
    run = run_agent(["bash", "-c", writer], tmp_path, 1, tmp_path)
    assert run["timed_out"] and run["exit_code"] is None and time.monotonic() - started < 15
    assert "process group was killed" in run["notes"][0]
    lines = (tmp_path / "late.txt").read_text().count("\n")
    time.sleep(1)
    assert (tmp_path / "late.txt").read_text().count("\n") == lines  # nothing keeps writing after the kill


@POSIX
def test_processes_left_behind_after_a_normal_exit_are_killed(tmp_path):
    run = run_agent(["bash", "-c", f"(sleep 1; echo late > {tmp_path}/late.txt) & echo done"], tmp_path, 30, tmp_path)
    assert run["exit_code"] == 0 and not run["timed_out"] and run["stdout"].strip() == "done"
    assert run["notes"] and "left processes running" in run["notes"][0]
    time.sleep(1.5)
    assert not (tmp_path / "late.txt").exists()


def test_grade_without_libreoffice_is_a_grader_error_not_a_fail(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(engine_mod, "_DEFAULT", None)
    monkeypatch.setattr(engine_mod, "find_soffice", lambda: None)
    case = load_case("xlsx-006")
    shutil.copytree(case.input_dir, tmp_path / "w")
    assert main(["grade", "xlsx-006", str(tmp_path / "w"), "--harbor-reward", str(tmp_path / "rw")]) == 4
    err = capsys.readouterr().err
    assert "LibreOffice was not found" in err and "Traceback" not in err
    assert json.loads((tmp_path / "rw" / "reward.json").read_text())["grader_error"] == 1


def test_grade_record_rejects_contradicting_run_status(tmp_path):
    code = main(["grade", "xlsx-006", str(tmp_path), "--record", str(tmp_path / "res"), "--agent", "a",
                 "--run-status", "completed", "--exit-code", "1"])
    assert code == 2 and not (tmp_path / "res").exists()


@needs_libreoffice
def test_grade_record_keeps_timeouts_and_hashes_prompt_and_inputs(tmp_path):
    work = tmp_path / "try"
    assert main(["prepare", "xlsx-004", str(work)]) == 0
    code = main(["grade", "xlsx-004", str(work), "--record", str(tmp_path / "res"), "--agent", "cursor",
                 "--model", "m", "--harness", "harbor", "--run-status", "timeout", "--timeout", "900"])
    assert code == 1
    (path,) = (tmp_path / "res" / "xlsx-004").glob("*.json")
    r = json.loads(path.read_text())
    assert r["run_status"] == "timeout" and r["timed_out"] is True
    assert r["classification"]["auto"] == "agent-or-environment-error"
    assert r["budget"]["timeout_sec"] == 900
    assert r["prompt_sha256"] == sha256_file(work / "TASK.md")
    case = load_case("xlsx-004")
    assert r["inputs"] == {n: sha256_file(case.input_dir / n) for n in case.inputs}
    assert all(r["inputs"].values()) and validate_result(r) == []
