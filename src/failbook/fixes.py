"""Fixes: reusable remedies (a skill or a prompt addition) that can be retested against cases.

A fix lives in ``fixes/<id>/`` with a ``fix.yaml`` and either a ``skill/``
directory (an Agent Skill with a SKILL.md) or a ``prompt.md`` that is appended
to the task. Retesting a case with ``--fix <id>`` shows whether the remedy works.
"""

from __future__ import annotations

import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Any

import yaml

from .case import CaseError, repo_root

if TYPE_CHECKING:  # pragma: no cover
    from .runner import AgentPreset


def fixes_dir() -> Path:
    root = repo_root()
    if root is None:
        raise CaseError("cannot find the fixes/ directory; run inside a Failbook checkout")
    return root / "fixes"


@dataclass
class Fix:
    dir: Path
    spec: dict[str, Any]

    @property
    def id(self) -> str:
        return self.spec["id"]

    @property
    def kind(self) -> str:
        return self.spec["kind"]

    @property
    def skill_dir(self) -> Path:
        return self.dir / "skill"

    @property
    def skill_name(self) -> str:
        text = (self.skill_dir / "SKILL.md").read_text(encoding="utf-8")
        for line in text.splitlines()[1:]:
            if line.startswith("name:"):
                return line.split(":", 1)[1].strip()
            if line.strip() == "---":
                break
        return self.id

    def apply(self, workdir: Path, prompt: str, preset: "AgentPreset | None") -> str:
        """Install the fix into ``workdir`` and return the (possibly extended) prompt."""
        if self.kind == "prompt":
            return prompt.rstrip() + "\n\n" + (self.dir / "prompt.md").read_text(encoding="utf-8").strip() + "\n"
        if self.kind == "skill":
            skills_dir = preset.skills_dir if preset else None
            if skills_dir:
                shutil.copytree(self.skill_dir, Path(workdir) / skills_dir / self.skill_name)
            else:
                # Agents without skill discovery get the skill text as project instructions.
                name = preset.instructions_file if preset else "AGENTS.md"
                body = (self.skill_dir / "SKILL.md").read_text(encoding="utf-8")
                (Path(workdir) / name).write_text(body, encoding="utf-8")
            return prompt
        raise CaseError(f"fix {self.id}: unknown kind {self.kind!r}")


def load_fix(ref: str) -> Fix:
    path = Path(ref)
    if not (path / "fix.yaml").is_file():
        path = fixes_dir() / ref
    spec_path = path / "fix.yaml"
    if not spec_path.is_file():
        raise CaseError(f"no fix named {ref!r} (looked for {spec_path})")
    spec = yaml.safe_load(spec_path.read_text(encoding="utf-8"))
    fix = Fix(path.resolve(), spec)
    if fix.kind == "skill" and not (fix.skill_dir / "SKILL.md").is_file():
        raise CaseError(f"fix {fix.id}: skill/SKILL.md is missing")
    if fix.kind == "prompt" and not (fix.dir / "prompt.md").is_file():
        raise CaseError(f"fix {fix.id}: prompt.md is missing")
    return fix


def all_fixes() -> list[Fix]:
    d = fixes_dir()
    if not d.is_dir():
        return []
    return [load_fix(str(p)) for p in sorted(d.iterdir()) if (p / "fix.yaml").is_file()]
