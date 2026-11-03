"""The audit script against its regression corpus (tests/audit_corpus/README.md).

Each probe is a before/after workbook pair from the reviews of the audit, labelled with the outcome a
correct audit gives. A probe whose outcome leaves its label fails, and so does a documented limitation
whose outcome changes: then the limitation was fixed or got worse, and probes.yaml and the skill's
SKILL.md must say which.
"""

import re
from pathlib import Path

import pytest
import yaml

from conftest import needs_libreoffice
from failbook import audit as audit_mod

CORPUS = Path(__file__).parent / "audit_corpus"
LABELS = {"error", "warning", "clean", "error|warning", "warning|clean", "unreadable", "any"}
#: The round that built a probe (docs/DECISIONS.md D9) and the id prefix that names it.
SOURCES = {"review-2": "review2-", "fix-1": "fix1-", "recheck-1": "recheck1-", "fix-2": "fix2-", "recheck-2": "recheck2-",
           "fix-3": "fix3-", "recheck-3": "recheck3-", "fix-4": "fix4-", "check": "check-"}
KEYS = {"id", "label", "edit", "why", "after", "before", "source", "require", "forbid", "known_limit", "now",
        "libreoffice_sample", "with_libreoffice"}


class _UniqueKeys(yaml.SafeLoader):
    """A YAML loader that refuses a key given twice in one mapping (PyYAML would keep the last one)."""


def _mapping(loader, node, deep=False):
    keys = [loader.construct_object(k, deep=deep) for k, _ in node.value]
    dup = {k for k in keys if keys.count(k) > 1}
    if dup:
        raise yaml.constructor.ConstructorError(None, None, f"duplicate key(s) {sorted(dup)}", node.start_mark)
    return yaml.SafeLoader.construct_mapping(loader, node, deep)


_UniqueKeys.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG,
                            lambda loader, node: _mapping(loader, node, deep=True))
PROBES = yaml.load((CORPUS / "probes.yaml").read_text(encoding="utf-8"), Loader=_UniqueKeys)


def _path(spec, tmp_path):
    if spec is None:
        return None
    if isinstance(spec, dict):  # a directory named like a workbook, which the audit must refuse
        path = tmp_path / spec["directory"]
        path.mkdir()
        return path
    return CORPUS / "workbooks" / spec


def _outcome(probe, tmp_path, recalc=False):
    """('error' | 'warning' | 'clean' | 'unreadable', findings as text)."""
    try:
        found = audit_mod.audit(_path(probe["after"], tmp_path), _path(probe.get("before"), tmp_path), recalc=recalc)
    except audit_mod.Unreadable as exc:
        return "unreadable", str(exc)
    found = [f for f in found if f.check != "recalculation-skipped"]
    outcome = "error" if any(f.level == "error" for f in found) else "warning" if found else "clean"
    return outcome, "\n".join(f"[{f.level}] {f.check} at {f.where}: {f.message}" for f in found)


def _check(probe, outcome, text, label):
    shown = f"{probe['id']}: {probe.get('edit', '')}\noutcome {outcome}, label {label}\n{text or 'no findings'}"
    if probe.get("known_limit") and probe.get("now") == "require-unmet":
        # The outcome is right but for the wrong cells: the finding the label is about is missing.
        if label != "any":
            assert outcome in label.split("|"), shown
        assert not re.search(probe["require"], text), (
            f"documented limitation fixed: a finding now matches {probe['require']!r}. Remove 'now' (and "
            f"'known_limit') from probes.yaml and update SKILL.md.\n{shown}")
        return
    if probe.get("known_limit") and "now" in probe:
        assert outcome == probe["now"], (
            f"documented limitation changed from {probe['now']!r} to {outcome!r}. If {outcome!r} is within the label "
            f"{label!r}, the limitation is fixed: remove 'now' (and 'known_limit') from probes.yaml and update SKILL.md; "
            f"otherwise it got worse.\n{shown}")
        return
    if label != "any":
        assert outcome in label.split("|"), shown
    if outcome != "unreadable":
        if probe.get("require"):
            assert re.search(probe["require"], text), f"no finding matches {probe['require']!r}\n{shown}"
        if probe.get("forbid"):
            assert not re.search(probe["forbid"], text), f"a finding matches {probe['forbid']!r}\n{shown}"


def test_probes_are_well_formed():
    """A misspelled key would silently switch a check off, so every entry is checked field by field."""
    ids = [p["id"] for p in PROBES]
    assert len(ids) == len(set(ids))
    used = set()
    for p in PROBES:
        pid = p.get("id")
        assert set(p) <= KEYS, (pid, sorted(set(p) - KEYS))
        assert p["label"] in LABELS, pid
        assert p.get("edit") and p.get("why") and p.get("after"), pid
        assert p.get("source") in SOURCES and pid.startswith(SOURCES[p["source"]]), (pid, p.get("source"))
        for key in ("require", "forbid"):
            if key in p:
                re.compile(p[key])
        for key in ("after", "before"):
            if isinstance(p.get(key), str):
                assert (CORPUS / "workbooks" / p[key]).is_file(), (pid, p[key])
                used.add(p[key])
        if "with_libreoffice" in p:
            assert p.get("libreoffice_sample") and p["with_libreoffice"] in LABELS, pid
        if "now" in p:
            assert p.get("known_limit"), pid
            if p["now"] == "require-unmet":
                assert p.get("require") and "forbid" not in p, pid
            else:  # require and forbid are not checked when the outcome is pinned
                assert p["now"] not in p["label"].split("|") and not {"require", "forbid"} & set(p), pid
    stored = {f.relative_to(CORPUS / "workbooks").as_posix() for f in (CORPUS / "workbooks").rglob("*") if f.is_file()}
    assert stored == used, f"workbooks no probe uses: {sorted(stored - used)}"


@pytest.mark.parametrize("probe", PROBES, ids=[p["id"] for p in PROBES])
def test_probe(probe, tmp_path):
    outcome, text = _outcome(probe, tmp_path)
    _check(probe, outcome, text, probe["label"])


LO_PROBES = [p for p in PROBES if p.get("libreoffice_sample")]


@needs_libreoffice
@pytest.mark.parametrize("probe", LO_PROBES, ids=[p["id"] for p in LO_PROBES])
def test_probe_with_libreoffice(probe, tmp_path):
    outcome, text = _outcome(probe, tmp_path, recalc=True)
    _check(probe, outcome, text, probe.get("with_libreoffice", probe["label"]))
