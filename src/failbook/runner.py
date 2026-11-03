"""Running an agent on cases and recording result files.

``failbook run`` is a convenience for local retests. It does not sandbox the
agent: the agent runs with your user account in a temporary working directory.
For isolated, comparable runs, export the cases to Harbor (``failbook export
harbor``) and run them in Docker. Both paths produce the same result schema.
"""

from __future__ import annotations

import datetime as dt
import json
import os
import re
import shlex
import shutil
import subprocess
import sys
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from . import __version__
from .case import Case, repo_root, resolve_cases
from .fixes import Fix, load_fix

RESULT_SCHEMA = "failbook.result/1"


@dataclass
class AgentPreset:
    name: str
    argv: list[str]
    model_args: list[str] = field(default_factory=list)
    version_cmd: list[str] = field(default_factory=list)
    skills_dir: str | None = None          # where project skills are discovered, if supported
    instructions_file: str = "AGENTS.md"   # fallback for prompt-style fixes
    note: str = ""


PRESETS: dict[str, AgentPreset] = {
    "claude-code": AgentPreset(
        name="claude-code",
        argv=["claude", "-p", "{prompt}", "--dangerously-skip-permissions"],
        model_args=["--model", "{model}"],
        version_cmd=["claude", "--version"],
        skills_dir=".claude/skills",
        instructions_file="CLAUDE.md",
        note=("Runs Claude Code headless with permission prompts disabled, inside the working directory. "
              "Claude Code refuses this mode as root outside a recognised sandbox; run as a normal user, "
              "or use the Harbor export for a container."),
    ),
    "codex": AgentPreset(
        name="codex",
        # --full-auto was removed in Codex 0.147.0; --sandbox workspace-write is its replacement.
        argv=["codex", "exec", "--sandbox", "workspace-write", "--skip-git-repo-check", "{prompt}"],
        model_args=["--model", "{model}"],
        version_cmd=["codex", "--version"],
        instructions_file="AGENTS.md",
        note="Runs Codex CLI non-interactively; Codex's own sandbox limits writes to the working directory.",
    ),
    "gemini-cli": AgentPreset(
        name="gemini-cli",
        # Without --skip-trust, Gemini CLI silently drops --yolo in an untrusted folder.
        argv=["gemini", "--yolo", "--skip-trust", "-p", "{prompt}"],
        model_args=["--model", "{model}"],
        version_cmd=["gemini", "--version"],
        instructions_file="GEMINI.md",
        note="Runs Gemini CLI non-interactively, auto-approving tool calls (--yolo --skip-trust).",
    ),
}


def _slug(text: str) -> str:
    return re.sub(r"[^a-z0-9.]+", "-", text.lower()).strip("-")[:40] or "unknown"


def _now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")


def agent_version(preset: AgentPreset | None) -> str:
    if not preset or not preset.version_cmd or not shutil.which(preset.version_cmd[0]):
        return ""
    try:
        out = subprocess.run(preset.version_cmd, capture_output=True, text=True, timeout=30)
        return (out.stdout or out.stderr).strip().splitlines()[0][:100]
    except Exception:
        return ""


def build_argv(preset: AgentPreset | None, template: str | None, *, prompt: str, prompt_file: Path,
               workdir: Path, model: str) -> list[str]:
    values = {"prompt": prompt, "prompt_file": str(prompt_file), "workdir": str(workdir), "model": model}
    if preset is not None:
        argv = [a.format(**values) for a in preset.argv]
        if model:
            extra = [a.format(**values) for a in preset.model_args]
            # Options go before a trailing positional prompt (e.g. `codex exec [OPTIONS] [PROMPT]`).
            if preset.argv[-1] == "{prompt}":
                argv[-1:-1] = extra
            else:
                argv += extra
        return argv
    assert template is not None
    # Substitute after splitting so that a multi-line prompt stays a single argument.
    return [part.format(**values) for part in shlex.split(template)]


def sha256_file(path: Path) -> str | None:
    import hashlib

    path = Path(path)
    if not path.is_file():
        return None
    return "sha256:" + hashlib.sha256(path.read_bytes()).hexdigest()


def sha256_text(text: str) -> str:
    import hashlib

    return "sha256:" + hashlib.sha256(text.encode("utf-8")).hexdigest()


def environment_info(isolation: str) -> dict[str, Any]:
    import platform

    return {
        "os": platform.platform(),
        "python": platform.python_version(),
        "isolation": isolation,
    }


def run_status(exit_code: int | None, timed_out: bool, ran: bool = True) -> str:
    if not ran:
        return "not-run"
    if timed_out:
        return "timeout"
    if exit_code not in (0, None):
        return "agent-error"
    return "completed"


def auto_classification(verdict: str, status: str) -> str:
    """A first, mechanical classification. Reviewers refine it (model vs tool failure) by hand."""
    if status in ("timeout", "agent-error"):
        return "agent-or-environment-error"
    if verdict == "inconclusive":
        return "evaluator-inconclusive"
    if verdict == "fail":
        return "task-failure"
    return "none"


def make_result(case: Case, report, *, seed: int, lang: str, agent: str, agent_ver: str, command: str,
                model: str, fix: str | None, harness: str, attempt: int, attempts: int,
                started_at: str, duration: float | None, exit_code: int | None, timed_out: bool,
                final_message: str, submitter: str, notes: str = "", inputs: dict | None = None,
                outputs: dict | None = None, prompt: str | None = None, timeout_sec: int | None = None,
                isolation: str = "unknown", ran: bool = True) -> dict[str, Any]:
    status = run_status(exit_code, timed_out, ran)
    return {
        "schema": RESULT_SCHEMA,
        "case": case.id,
        "case_version": case.version,
        "seed": seed,
        "lang": lang,
        "agent": {"name": agent, "version": agent_ver, "command": command},
        "model": model,
        "fix": fix,
        "harness": harness,
        "attempt": attempt,
        "attempts": attempts,
        "started_at": started_at,
        "duration_sec": None if duration is None else round(duration, 1),
        "exit_code": exit_code,
        "timed_out": timed_out,
        "run_status": status,
        "verdict": report.verdict,
        "score": round(report.score, 4),
        "rules": [r.to_dict() for r in report.rules],
        "classification": {"auto": auto_classification(report.verdict, status), "reviewed": None, "note": ""},
        "engine": report.engine,
        "failbook_version": __version__,
        "prompt_sha256": sha256_text(prompt) if prompt is not None else None,
        "inputs": inputs or {},
        "outputs": outputs or {},
        "environment": environment_info(isolation),
        "budget": {"timeout_sec": timeout_sec, "cost_usd": None, "tokens": None},
        "final_message": final_message[-4000:],
        "claimed_done": None,
        "submitter": submitter,
        "notes": notes,
    }


def write_result(result: dict[str, Any], results_dir: Path) -> Path:
    day = result["started_at"][:10]
    name = f"{day}-{_slug(result['agent']['name'])}-{_slug(result['model'] or 'default')}"
    if result.get("fix"):
        name += f"-fix-{_slug(result['fix'])}"
    name += f"-s{result['seed']}-{uuid.uuid4().hex[:6]}.json"
    path = Path(results_dir) / result["case"] / name
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return path


def run_one(case: Case, *, preset: AgentPreset | None, template: str | None, agent_name: str, model: str,
            seed: int, lang: str, fix: Fix | None, timeout: int, run_dir: Path, dry_run: bool):
    from .grade import grade, prepare

    work = run_dir / "work"
    meta = prepare(case, work, seed, lang=lang, write_task=False)
    inputs = {name: sha256_file(work / name) for name in case.inputs}
    prompt = case.prompt(meta, lang)
    if fix is not None:
        prompt = fix.apply(work, prompt, preset)
    prompt_file = run_dir / "TASK.md"
    prompt_file.write_text(prompt, encoding="utf-8")
    argv = build_argv(preset, template, prompt=prompt, prompt_file=prompt_file, workdir=work, model=model)
    shown = " ".join(shlex.quote(a if a != prompt else "<TASK.md>") for a in argv)
    if dry_run:
        print(f"[dry-run] cd {work} && {shown}")
        return None, {"command": shown}
    log = run_dir / "agent.log"
    started = time.monotonic()
    timed_out = False
    try:
        proc = subprocess.run(argv, cwd=work, capture_output=True, text=True, timeout=timeout)
        stdout, stderr, code = proc.stdout, proc.stderr, proc.returncode
    except subprocess.TimeoutExpired as exc:
        timed_out = True
        stdout = exc.stdout.decode() if isinstance(exc.stdout, bytes) else (exc.stdout or "")
        stderr = exc.stderr.decode() if isinstance(exc.stderr, bytes) else (exc.stderr or "")
        code = None
    except FileNotFoundError:
        raise SystemExit(f"error: agent command not found: {argv[0]}")
    duration = time.monotonic() - started
    log.write_text(f"$ {shown}\n\n--- stdout ---\n{stdout}\n--- stderr ---\n{stderr}\n", encoding="utf-8")
    report = grade(case, work, seed)
    (run_dir / "grade.json").write_text(json.dumps(report.to_dict(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    info = {
        "command": shown, "exit_code": code, "timed_out": timed_out, "final_message": stdout.strip(),
        "duration": duration, "inputs": inputs, "prompt": prompt,
        "outputs": {name: sha256_file(work / name) for name in case.outputs},
    }
    return report, info


def run_cases(args) -> int:
    if args.agent == "custom":
        if not args.cmd:
            print("error: --agent custom needs --cmd", file=sys.stderr)
            return 2
        preset, template, agent_name = None, args.cmd, "custom"
    elif args.agent in PRESETS:
        preset, template, agent_name = PRESETS[args.agent], None, args.agent
        if not args.dry_run and not shutil.which(preset.argv[0]):
            print(f"error: `{preset.argv[0]}` is not on PATH; install {args.agent} or use --agent custom", file=sys.stderr)
            return 2
    else:
        print(f"error: unknown agent {args.agent!r}; presets: {', '.join(PRESETS)} or 'custom'", file=sys.stderr)
        return 2
    fix = load_fix(args.fix) if args.fix else None
    version = agent_version(preset)
    cases = resolve_cases(args.cases or None)
    runs_root = Path(args.runs_dir)
    results_dir = Path(args.results)
    if not results_dir.is_absolute() and repo_root() and not Path.cwd().joinpath(results_dir).exists():
        results_dir = repo_root() / results_dir
    summary = []
    for case in cases:
        for attempt in range(1, args.attempts + 1):
            stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
            run_dir = runs_root / f"{stamp}-{_slug(agent_name)}-{case.id}-a{attempt}"
            run_dir.mkdir(parents=True, exist_ok=True)
            started_at = _now()
            report, info = run_one(
                case, preset=preset, template=template, agent_name=agent_name, model=args.model,
                seed=args.seed, lang=args.lang, fix=fix, timeout=args.timeout, run_dir=run_dir,
                dry_run=args.dry_run,
            )
            if report is None:
                continue
            result = make_result(
                case, report, seed=args.seed, lang=args.lang, agent=agent_name, agent_ver=version,
                command=info["command"], model=args.model, fix=fix.id if fix else None, harness="failbook-run",
                attempt=attempt, attempts=args.attempts, started_at=started_at, duration=info["duration"],
                exit_code=info["exit_code"], timed_out=info["timed_out"], final_message=info["final_message"],
                submitter=args.submitter, inputs=info["inputs"], outputs=info["outputs"], prompt=info["prompt"],
                timeout_sec=args.timeout, isolation="none (local process, temporary directory)",
            )
            (run_dir / "result.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            path = write_result(result, results_dir)
            summary.append((case.id, attempt, report.verdict))
            print(f"{case.id} attempt {attempt}: {report.verdict.upper()} (run {result['run_status']})  -> {path}")
            for r in report.rules:
                if r.status != "pass":
                    print(f"    [{r.status}] {r.id}: {r.detail}")
    if summary:
        passed = sum(v == "pass" for _, _, v in summary)
        print(f"\n{passed}/{len(summary)} attempts passed. Working directories and logs are in {runs_root}/")
    return 0
