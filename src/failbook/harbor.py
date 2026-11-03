"""Export cases as Harbor tasks, and fixes as SkillEvaluator skills with native Harbor evals.

Harbor (https://github.com/harbor-framework/harbor) runs agents in sandboxed
containers and is the execution layer behind NVIDIA SkillEvaluator's live
evaluation. Exporting to it means Failbook does not need its own sandboxing or
agent adapters: any agent Harbor supports can be retested on any case.

Layout of one exported task::

    <case-id>/
      task.toml
      instruction.md            the prompt (paths point at /workspace)
      environment/Dockerfile    Python, openpyxl, pandas, LibreOffice; inputs copied to /workspace
      environment/workspace/    the input files
      tests/test.sh             runs the Failbook grader; writes /logs/verifier/reward.json
      tests/lib/failbook/       the grader (uploaded only after the agent has finished; run in
                                isolated mode so files the agent left in /workspace are never imported)
      tests/cases/<case>/       the case definition, without its solutions
      solution/solve.sh         applies the reference solution (for Harbor's oracle agent)
"""

from __future__ import annotations

import json
import shutil
import tomllib
from pathlib import Path

from .case import Case

DOCKERFILE = """FROM python:3.13-slim-trixie

# Same toolbox for every agent: Python with openpyxl/pandas, and LibreOffice for recalculation.
RUN apt-get -o Acquire::Retries=3 update && \\
    apt-get -o Acquire::Retries=3 install -y --no-install-recommends \\
        libreoffice-calc-nogui bash curl git ca-certificates && \\
    rm -rf /var/lib/apt/lists/*
RUN pip install --no-cache-dir "openpyxl>=3.1,<4" "PyYAML>=6" "pandas>=2"

RUN mkdir -p /workspace /logs/agent /logs/verifier
COPY workspace/ /workspace/
WORKDIR /workspace
"""

#: Runs the packaged grader without letting anything else onto its import path: ``-I`` ignores the
#: working directory, PYTHON* variables and user site-packages, and only ``<dir>/lib`` is added.
_RUN_FAILBOOK = (
    "python3 -I -c 'import sys; sys.path.insert(0, sys.argv.pop(1)); "
    "from failbook.cli import main; sys.exit(main(sys.argv[1:]))'"
)

TEST_SH = """#!/bin/bash
# Grades /workspace with Failbook and writes Harbor rewards. Never fails the script itself:
# a crash in the grader must not be confused with a failing agent; it is reported as grader_error 1.
set -uo pipefail
tests_dir="${{HARBOR_TESTS_DIR:-/tests}}"
logs_dir="${{HARBOR_LOGS_DIR:-/logs}}"
verifier="$logs_dir/verifier"
mkdir -p "$verifier"
# The verifier directory is writable while the agent works: never keep a reward written before grading.
rm -f "$verifier/reward.json" "$verifier/reward.txt" "$verifier/failbook.json"
# Run from a neutral directory, in isolated mode: files in /workspace (the agent's) are never imported.
cd "$tests_dir" 2>/dev/null || cd /
export FAILBOOK_CASES="$tests_dir/cases"
{run} "$tests_dir/lib" grade {case} /workspace --seed {seed} \\
    --json "$verifier/failbook.json" --harbor-reward "$verifier"
status=$?
if [ ! -s "$verifier/reward.json" ]; then
    echo '{grader_error_json}' > "$verifier/reward.json"
    echo 0 > "$verifier/reward.txt"
fi
echo "failbook grade exited with $status (0 pass, 1 fail, 3 inconclusive, 4 grader error)"
exit 0
"""

SOLVE_SH = """#!/bin/bash
set -euo pipefail
here="$(cd "$(dirname "${{BASH_SOURCE[0]}}")" && pwd)"
cd "$here"
export FAILBOOK_CASES="$here/cases"
{run} "$here/lib" solve {case} /workspace --kind reference --seed {seed}
"""

#: Every reward.json has the same keys: Harbor averages each key over trials and counts a missing key as 0.
GRADER_ERROR_REWARD = {"reward": 0.0, "overall": 0.0, "rules_passed": 0.0, "inconclusive": 0, "grader_error": 1}


def write_reward(report, outdir: Path) -> None:
    """Harbor reward contract: reward.txt plus a reward.json whose top-level values are all numbers.

    ``reward``/``overall`` is 1 only for a pass. An inconclusive verdict (the grader could not decide)
    also gives 0, because Harbor accepts numbers only; ``inconclusive`` = 1 marks those trials so they
    can be left out (see the README written by ``export_harbor``).
    """
    reward = 1.0 if report.verdict == "pass" else 0.0
    _write_reward_files(Path(outdir), {
        "reward": reward,
        "overall": reward,
        "rules_passed": round(report.score, 4),
        "inconclusive": 1 if report.verdict == "inconclusive" else 0,
        "grader_error": 0,
    })


def write_grader_error(outdir: Path) -> None:
    """Rewards for a trial whose grading crashed: no verdict, marked with ``grader_error`` = 1."""
    _write_reward_files(Path(outdir), dict(GRADER_ERROR_REWARD))


def _write_reward_files(outdir: Path, data: dict) -> None:
    outdir.mkdir(parents=True, exist_ok=True)
    (outdir / "reward.json").write_text(json.dumps(data) + "\n", encoding="utf-8")
    (outdir / "reward.txt").write_text(f"{data['reward']:g}\n", encoding="utf-8")


def _package_src() -> Path:
    return Path(__file__).resolve().parent


def _copy_lib(dest: Path) -> None:
    shutil.copytree(_package_src(), dest / "failbook", ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))


def _copy_case(case: Case, dest: Path, *, with_reference: bool) -> None:
    target = dest / case.dir.name
    target.mkdir(parents=True)
    for name in ("case.yaml", "case.py", "task.md", "task.zh.md"):
        if (case.dir / name).is_file():
            shutil.copyfile(case.dir / name, target / name)
    if (case.input_dir).is_dir():
        shutil.copytree(case.input_dir, target / "input")
    if with_reference:
        shutil.copyfile(case.dir / "reference.py", target / "reference.py")


def _toml_str(s: str) -> str:
    return json.dumps(s, ensure_ascii=False)


def task_name(case: Case, seed: int) -> str:
    return f"failbook/{case.id}" + (f"-seed{seed}" if seed else "")


def export_task(case: Case, out: Path, *, seed: int = 0, lang: str = "en", skills_dir: str | None = None) -> Path:
    from .grade import prepare

    task_dir = Path(out) / (case.id + (f"-seed{seed}" if seed else ""))
    if task_dir.exists():
        shutil.rmtree(task_dir)
    env = task_dir / "environment"
    workspace = env / "workspace"
    meta = prepare(case, workspace, seed, lang=lang, write_task=False)
    (env / "Dockerfile").write_text(DOCKERFILE, encoding="utf-8")
    (task_dir / "instruction.md").write_text(case.prompt(meta, lang, where="/workspace"), encoding="utf-8")

    tests = task_dir / "tests"
    _copy_lib(tests / "lib")
    _copy_case(case, tests / "cases", with_reference=False)
    test_sh = tests / "test.sh"
    test_sh.write_text(TEST_SH.format(case=case.id, seed=seed, run=_RUN_FAILBOOK,
                                      grader_error_json=json.dumps(GRADER_ERROR_REWARD)),
                       encoding="utf-8", newline="\n")
    test_sh.chmod(0o755)

    solution = task_dir / "solution"
    _copy_lib(solution / "lib")
    _copy_case(case, solution / "cases", with_reference=True)
    solve_sh = solution / "solve.sh"
    solve_sh.write_text(SOLVE_SH.format(case=case.id, seed=seed, run=_RUN_FAILBOOK), encoding="utf-8", newline="\n")
    solve_sh.chmod(0o755)

    keywords = ", ".join(_toml_str(t) for t in ["failbook", case.spec.get("domain", "spreadsheet"), *case.spec.get("tags", [])])
    env_extra = f"skills_dir = {_toml_str(skills_dir)}\n" if skills_dir else ""
    toml = f"""schema_version = "1.4"

[task]
name = {_toml_str(task_name(case, seed))}
description = {_toml_str(case.title)}
keywords = [{keywords}]

[metadata]
entry_id = {_toml_str(case.id + (f"-seed{seed}" if seed else ""))}
failbook_case = {_toml_str(case.id)}
failbook_case_version = {case.version}
failbook_seed = {seed}
failbook_lang = {_toml_str(lang)}

[agent]
timeout_sec = 900.0

[verifier]
timeout_sec = 300.0

[environment]
cpus = 2
memory_mb = 4096
storage_mb = 4096
network_mode = "public"
workdir = "/workspace"
{env_extra}"""
    tomllib.loads(toml)
    (task_dir / "task.toml").write_text(toml, encoding="utf-8")
    return task_dir


HARBOR_README = """# Failbook cases as Harbor tasks

Generated by `failbook export harbor`. Run them with Harbor, for example:

```bash
harbor run -p . -a claude-code -m <model> -k 3
```

Each trial writes `/logs/verifier/failbook.json` with the per-rule verdicts, and a `reward.json` with
these numeric keys:

| Key | Meaning |
|---|---|
| `reward`, `overall` | 1 only when every rule passes, otherwise 0 |
| `rules_passed` | share of the decided rules that passed |
| `inconclusive` | 1 when the grader could not decide (for example a function LibreOffice lacks) |
| `grader_error` | 1 when the grader itself crashed, so there is no verdict |

Harbor accepts only numbers, so inconclusive and crashed trials also get `reward` 0. They are not failures
of the agent. Harbor's mean reward (and, for a SkillEvaluator export, pass@k and the lift computed from
`overall`) counts them as 0, so read the means of `inconclusive` and `grader_error` next to it, and leave
those trials out when you compute a pass rate. With [jq](https://jqlang.org/), over one job:

```bash
jq -s '[.[] | .verifier_result.rewards // empty | select(.inconclusive == 0 and .grader_error == 0)]
       | {decided: length, passed: (map(.reward) | add // 0)}' jobs/<job-name>/*/result.json
```

Trials without a `verifier_result` (for example a verifier timeout) are left out by `// empty`; Harbor
itself counts them as 0 as well.
"""


def export_harbor(cases: list[Case], out: Path, *, seed: int = 0, lang: str = "en") -> list[Path]:
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    paths = [export_task(c, out, seed=seed, lang=lang) for c in cases]
    (out / "README.md").write_text(HARBOR_README, encoding="utf-8")
    return paths


def export_skillevaluator(cases: list[Case], out: Path, *, fix: str, seed: int = 0, lang: str = "en") -> Path:
    """Package a skill fix with the cases as native Harbor evals, for SkillEvaluator Tier 3.

    ``skillevaluator tier3 evaluate <dir>`` then runs every case with and without
    the skill and reports the lift.
    """
    from .fixes import SKILL_COPY_IGNORE, load_fix

    f = load_fix(fix)
    if f.kind != "skill":
        raise ValueError(f"fix {f.id} is a {f.kind} fix; SkillEvaluator evaluates skills only")
    skill_root = Path(out) / f.skill_name
    if skill_root.exists():
        shutil.rmtree(skill_root)
    shutil.copytree(f.skill_dir, skill_root, ignore=SKILL_COPY_IGNORE)
    harbor_dir = skill_root / "evals" / "harbor"
    harbor_dir.mkdir(parents=True)
    tasks = [export_task(c, harbor_dir, seed=seed, lang=lang, skills_dir="/workspace/skills") for c in cases]
    lines = ["[dataset]", 'name = "failbook-spreadsheet"',
             'description = "Failbook spreadsheet cases (native Harbor tasks)"', ""]
    for c in cases:
        lines += ["[[tasks]]", f"name = {_toml_str(task_name(c, seed))}", ""]
    (harbor_dir / "dataset.toml").write_text("\n".join(lines), encoding="utf-8")
    (skill_root / "evals" / "config.yml").write_text(
        "schema_version: 1\n\n"
        "harbor:\n"
        "  task_source: native_harbor\n"
        "  n_attempts: 3\n"
        "  # Rewards are 0 or 1, so an attempt counts as passed only when every rule passed.\n"
        "  # Inconclusive and grader-error attempts also get overall 0 (Harbor accepts numbers only),\n"
        "  # so pass@k and the lift count them as not passed. Their share is reported as the custom\n"
        "  # metrics `inconclusive` and `grader_error`; check both arms before reading the lift.\n"
        "  pass_threshold: 1.0\n\n"
        "grading:\n"
        "  mode: custom_only\n",
        encoding="utf-8",
    )
    for t in tasks:
        (t / "environment" / "workspace" / "skills").mkdir(exist_ok=True)
    return skill_root
