"""Loading cases from ``cases/<id>-<slug>/``.

A case directory contains:

``case.yaml``      metadata: id, titles, the failure story, provenance, expected failing rules
``task.md``        the task given to the agent (English); ``task.zh.md`` is the Chinese version
``case.py``        ``params(seed)``, ``build(meta, outdir)``, optional ``scenarios(ctx)`` and the rules
``reference.py``   ``solve(workdir, meta)``: a correct solution, used to prove the rules accept good work
``failure.py``     ``solve(workdir, meta)``: reproduces the recorded failure, proving the rules catch it
``alt_*.py``       optional: other valid solutions, which must pass as well
``input/``         committed inputs for the canonical seed 0
``failure/``       committed failure artifact for seed 0
"""

from __future__ import annotations

import importlib.util
import os
import re
import shutil
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from types import ModuleType
from typing import Any

import yaml

from .checks import RuleSpec

TASK_FOOTER = {
    "en": (
        "\n---\n"
        "Work on the file(s) in {where} and save your changes in place, "
        "keeping the same file name(s): {files}. "
        "Do not change anything the task does not ask you to change.\n"
    ),
    "zh": (
        "\n---\n"
        "请直接处理{where_zh}下的文件，并以原文件名原地保存：{files}。"
        "任务没有要求修改的内容，请保持不变。\n"
    ),
}

_PLACEHOLDER = re.compile(r"\{\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\}\}")


class CaseError(Exception):
    pass


def repo_root() -> Path | None:
    env = os.environ.get("FAILBOOK_ROOT")
    if env:
        return Path(env).resolve()
    here = Path.cwd().resolve()
    for candidate in (here, *here.parents):
        if (candidate / "cases").is_dir() and (candidate / "pyproject.toml").is_file():
            return candidate
    pkg_root = Path(__file__).resolve().parents[2]
    if (pkg_root / "cases").is_dir():
        return pkg_root
    return None


def cases_dir() -> Path:
    env = os.environ.get("FAILBOOK_CASES")
    if env:
        return Path(env).resolve()
    root = repo_root()
    if root is None:
        raise CaseError(
            "cannot find the cases/ directory. Run inside a Failbook checkout or set FAILBOOK_CASES."
        )
    return root / "cases"


def _import(path: Path, name: str) -> ModuleType:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:  # pragma: no cover
        raise CaseError(f"cannot import {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


@dataclass
class Case:
    dir: Path
    spec: dict[str, Any]
    module: ModuleType
    rules: list[RuleSpec] = field(default_factory=list)

    # -- metadata -------------------------------------------------------------
    @property
    def id(self) -> str:
        return self.spec["id"]

    @property
    def slug(self) -> str:
        return self.spec["slug"]

    @property
    def version(self) -> int:
        return int(self.spec.get("version", 1))

    @property
    def title(self) -> str:
        return self.spec["title"]

    @property
    def title_zh(self) -> str:
        return self.spec.get("title_zh") or self.title

    @property
    def inputs(self) -> list[str]:
        return list(self.spec["files"]["inputs"])

    @property
    def outputs(self) -> list[str]:
        return list(self.spec["files"]["outputs"])

    @property
    def expected_failures(self) -> list[str]:
        return list(self.spec.get("failure", {}).get("failing_rules", []))

    @property
    def input_dir(self) -> Path:
        return self.dir / "input"

    @property
    def failure_dir(self) -> Path:
        return self.dir / "failure"

    def text(self, key: str, lang: str = "en") -> str:
        """Bilingual field lookup: ``failure.observed`` -> ``failure.observed_zh`` for zh."""
        node: Any = self.spec
        parts = key.split(".")
        for part in parts[:-1]:
            node = node.get(part, {})
        last = parts[-1]
        if lang == "zh" and node.get(f"{last}_zh"):
            return _join_cjk(str(node[f"{last}_zh"]).strip())
        return str(node.get(last, "")).strip()

    # -- generation -------------------------------------------------------------
    def params(self, seed: int = 0) -> dict[str, Any]:
        return self.module.params(seed)

    def build(self, seed: int, outdir: Path, *, engine=None) -> dict[str, Any]:
        """Generate the inputs for ``seed`` into ``outdir``.

        Inputs are written with openpyxl and then recalculated and re-saved by the
        spreadsheet engine, so they carry cached values like a file saved by a
        real spreadsheet application.
        """
        from .engine import default_engine

        meta = self.params(seed)
        outdir = Path(outdir)
        outdir.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix="failbook-build-") as raw:
            raw_dir = Path(raw)
            self.module.build(meta, raw_dir)
            to_finalize = []
            for name in self.inputs:
                src = raw_dir / name
                if not src.is_file():
                    raise CaseError(f"{self.id}: build() did not create {name}")
                if name.lower().endswith(".xlsx") and self.spec.get("finalize", True):
                    to_finalize.append(src)
                else:
                    shutil.copyfile(src, outdir / name)
            if to_finalize:
                engine = engine or default_engine()
                mapping = engine.recalc(to_finalize, raw_dir / "_final")
                for src, final in mapping.items():
                    shutil.copyfile(final, outdir / src.name)
            post = getattr(self.module, "post_build", None)
            if post is not None:
                post(meta, outdir)
        return meta

    def inputs_for(self, seed: int, outdir: Path, *, engine=None) -> dict[str, Any]:
        """Copy committed inputs for seed 0, otherwise build fresh ones."""
        outdir = Path(outdir)
        outdir.mkdir(parents=True, exist_ok=True)
        if seed == 0 and all((self.input_dir / n).is_file() for n in self.inputs):
            for name in self.inputs:
                shutil.copyfile(self.input_dir / name, outdir / name)
            return self.params(0)
        return self.build(seed, outdir, engine=engine)

    # -- task -----------------------------------------------------------------------
    def task(self, meta: dict[str, Any], lang: str = "en") -> str:
        fname = "task.zh.md" if lang == "zh" else "task.md"
        path = self.dir / fname
        if not path.is_file():
            path = self.dir / "task.md"
        template = path.read_text(encoding="utf-8")

        def sub(m: re.Match) -> str:
            key = m.group(1)
            if key not in meta:
                raise CaseError(f"{self.id}: task placeholder {{{{{key}}}}} missing from params()")
            return str(meta[key])

        return _PLACEHOLDER.sub(sub, template).strip() + "\n"

    def prompt(self, meta: dict[str, Any], lang: str = "en", where: str | None = None) -> str:
        """The full prompt: the task plus a footer saying where the files are."""
        files = ", ".join(f"`{f}`" for f in self.outputs)
        footer = TASK_FOOTER["zh" if lang == "zh" else "en"]
        return self.task(meta, lang) + footer.format(
            files=files,
            where=f"`{where}`" if where else "the current directory",
            where_zh=f" `{where}` " if where else "当前目录",
        )

    # -- solutions ------------------------------------------------------------------------
    def alternatives(self) -> list[str]:
        """Names of alternative valid solutions (``alt_*.py``)."""
        return sorted(p.stem for p in self.dir.glob("alt_*.py"))

    def solution(self, kind: str) -> ModuleType:
        if kind not in ("reference", "failure") and not kind.startswith("alt_"):
            raise CaseError(f"unknown solution kind {kind!r}")
        path = self.dir / f"{kind}.py"
        if not path.is_file():
            raise CaseError(f"{self.id}: {kind}.py is missing")
        return _import(path, f"failbook_cases.{_modname(self.dir.name)}.{kind}")

    def solve(self, kind: str, workdir: Path, meta: dict[str, Any]) -> None:
        self.solution(kind).solve(Path(workdir), meta)


_CJK = r"\u3000-\u303f\u4e00-\u9fff\uff00-\uffef"
_CJK_PUNCT = r"\u3000-\u303f\uff00-\uff0f\uff1a-\uff20"
_CJK_GAP = re.compile(rf"(?<=[{_CJK}])\s+(?=[{_CJK}])|(?<=[{_CJK_PUNCT}])\s+|\s+(?=[{_CJK_PUNCT}])")


def _join_cjk(text: str) -> str:
    """YAML folding turns line breaks into spaces; drop those between CJK characters."""
    return _CJK_GAP.sub("", text)


def _modname(dirname: str) -> str:
    return re.sub(r"\W", "_", dirname)


def load_case_dir(path: Path) -> Case:
    path = Path(path).resolve()
    spec_path = path / "case.yaml"
    if not spec_path.is_file():
        raise CaseError(f"{path} has no case.yaml")
    spec = yaml.safe_load(spec_path.read_text(encoding="utf-8"))
    for key in ("id", "slug", "title", "files"):
        if key not in spec:
            raise CaseError(f"{spec_path}: missing required key {key!r}")
    module = _import(path / "case.py", f"failbook_cases.{_modname(path.name)}.case")
    rules = [
        obj.__failbook_rule__
        for obj in vars(module).values()
        if callable(obj) and hasattr(obj, "__failbook_rule__")
    ]
    ids = [r.id for r in rules]
    if len(set(ids)) != len(ids):
        raise CaseError(f"{spec['id']}: duplicate rule ids")
    return Case(dir=path, spec=spec, module=module, rules=rules)


def all_cases(root: Path | None = None) -> list[Case]:
    root = Path(root) if root else cases_dir()
    cases = [load_case_dir(p) for p in sorted(root.iterdir()) if (p / "case.yaml").is_file()]
    return cases


def load_case(ref: str, root: Path | None = None) -> Case:
    """Find a case by id (``xlsx-001``), directory name, slug, or path."""
    p = Path(ref)
    if (p / "case.yaml").is_file():
        return load_case_dir(p)
    root = Path(root) if root else cases_dir()
    for d in sorted(root.iterdir()):
        if not (d / "case.yaml").is_file():
            continue
        if d.name == ref or d.name.startswith(ref + "-") or d.name.endswith("-" + ref):
            return load_case_dir(d)
    raise CaseError(f"no case matches {ref!r} in {root}")


def resolve_cases(refs: list[str] | None) -> list[Case]:
    if not refs or refs == ["all"]:
        return all_cases()
    return [load_case(r) for r in refs]
