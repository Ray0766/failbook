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


#: Where a skill is installed, relative to the working directory, for agents without skill discovery.
FALLBACK_SKILLS_DIR = "skills"
#: How SKILL.md refers to its own folder; replaced by the real path where Failbook knows it.
SKILL_DIR_PLACEHOLDER = "<this-skill-dir>"
#: Local build products that must not reach the agent's working directory.
SKILL_COPY_IGNORE = shutil.ignore_patterns("__pycache__", "*.pyc", ".DS_Store")


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
                shutil.copytree(self.skill_dir, Path(workdir) / skills_dir / self.skill_name, ignore=SKILL_COPY_IGNORE)
            else:
                # Agents without skill discovery get the whole skill folder (its scripts too) in
                # the working directory, and its text as project instructions pointing there.
                rel = f"{FALLBACK_SKILLS_DIR}/{self.skill_name}"
                shutil.copytree(self.skill_dir, Path(workdir) / rel, ignore=SKILL_COPY_IGNORE)
                name = preset.instructions_file if preset else "AGENTS.md"
                (Path(workdir) / name).write_text(self.instructions(rel), encoding="utf-8")
            return prompt
        raise CaseError(f"fix {self.id}: unknown kind {self.kind!r}")

    def instructions(self, rel: str) -> str:
        """SKILL.md as an instructions file for an agent that reads the skill folder at ``rel``."""
        body = (self.skill_dir / "SKILL.md").read_text(encoding="utf-8").replace(SKILL_DIR_PLACEHOLDER, rel)
        note = (f"> This is the `{self.skill_name}` skill, installed in `{rel}/` (relative to this file). "
                f"Paths such as `{rel}/scripts/...` below point into that folder.\n\n")
        if body.startswith("---\n"):
            end = body.find("\n---\n", 4)
            if end != -1:
                head, rest = body[:end + 5], body[end + 5:]
                return head + "\n" + note + rest.lstrip("\n")
        return note + body


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
