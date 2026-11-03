import shutil
import tomllib

from failbook.audit import audit
from failbook.case import load_case
from failbook.harbor import export_harbor


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
