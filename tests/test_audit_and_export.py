import shlex
import shutil
import subprocess
import sys
import tomllib

from failbook import __version__
from failbook.audit import audit
from failbook.case import load_case
from failbook.harbor import _RUN_FAILBOOK, export_harbor


def _errors(findings):
    return {f.check for f in findings if f.level == "error"}


def test_audit_flags_failures_and_is_quiet_on_references(tmp_path):
    for cid, expected in (("xlsx-001", "range-omits-adjacent-cells"), ("xlsx-004", "formula-moved-as-text"),
                          ("xlsx-012", "sheet-removed")):
        case = load_case(cid)
        name = case.outputs[0]
        before = case.input_dir / name
        assert expected in _errors(audit(case.failure_dir / name, before, recalc=False))
        work = tmp_path / cid
        shutil.copytree(case.input_dir, work)
        case.solve("reference", work, case.params(0))
        assert not _errors(audit(work / name, before, recalc=False)), f"false alarm on {cid} reference"


def test_harbor_export_layout(tmp_path):
    case = load_case("xlsx-002")
    export_harbor([case], tmp_path)
    task = tmp_path / "xlsx-002"
    config = tomllib.loads((task / "task.toml").read_text())
    assert config["schema_version"] == "1.4" and config["task"]["name"] == "failbook/xlsx-002"
    assert (task / "environment" / "workspace" / "pricing.xlsx").is_file()
    assert (task / "tests" / "test.sh").stat().st_mode & 0o111
    # The agent's container must not contain the rules or the solutions.
    assert not list((task / "environment").rglob("*.py"))
    assert not (task / "tests" / "cases" / case.dir.name / "reference.py").exists()
    assert (task / "solution" / "cases" / case.dir.name / "reference.py").exists()
    assert "/workspace" in (task / "instruction.md").read_text()
    test_sh = (task / "tests" / "test.sh").read_text()
    assert "python3 -I " in test_sh and "python3 -m" not in test_sh
    assert 'rm -f "$verifier/reward.json"' in test_sh  # a reward left by the agent is never kept
    assert "grader_error" in (tmp_path / "README.md").read_text()


def test_harbor_grader_never_imports_files_from_the_workspace(tmp_path):
    """Files an agent leaves in /workspace must not shadow the grader or the standard library."""
    case = load_case("xlsx-002")
    export_harbor([case], tmp_path / "out")
    lib = tmp_path / "out" / "xlsx-002" / "tests" / "lib"
    ws = tmp_path / "workspace"
    (ws / "failbook").mkdir(parents=True)
    (ws / "failbook" / "__init__.py").write_text("raise SystemExit('planted failbook package imported')\n")
    (ws / "inspect.py").write_text("raise SystemExit('planted inspect.py imported')\n")
    runner = _RUN_FAILBOOK.replace("python3", sys.executable, 1)
    proc = subprocess.run(f"{runner} {shlex.quote(str(lib))} --version", shell=True, cwd=ws, capture_output=True,
                          text=True, env={"PYTHONPATH": str(ws), "PATH": "/usr/bin:/bin"})
    assert proc.returncode == 0, proc.stderr
    assert proc.stdout.strip() == f"failbook {__version__}"

