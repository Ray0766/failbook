"""Running an agent on cases and recording result files.

``failbook run`` is a convenience for local retests. It does not sandbox the
agent: the agent runs unsandboxed as your user and can read (and write) the rest
of the file system. Its working directory is created outside the Failbook
checkout, so the rules and reference solutions are not next to it, but they are
not hidden from it either. For isolated, comparable runs, export the cases to
Harbor (``failbook export harbor``) and run them in Docker. Both paths produce
the same result schema.
"""

from __future__ import annotations

import datetime as dt
import json
import os
import re
import shlex
import shutil
import signal
import subprocess
import sys
import tempfile
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from . import __version__
from .case import Case, CaseError, cases_dir, repo_root, resolve_cases
from .fixes import Fix, load_fix

RESULT_SCHEMA = "failbook.result/1"
RUN_STATUSES = ("completed", "timeout", "agent-error", "not-run")

#: Recorded in every result of ``failbook run`` (environment.isolation).
LOCAL_ISOLATION = ("none: unsandboxed local process running as the current user; it could read and write "
                   "the rest of the file system. Working directory: a new directory outside the Failbook checkout")

#: Marks where a preset's model options go (removed when no model is given).
MODEL_ARGS = "{model_args}"


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
        argv=["claude", "-p", "{prompt}", "--dangerously-skip-permissions", MODEL_ARGS],
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
        argv=["codex", "exec", "--sandbox", "workspace-write", "--skip-git-repo-check", MODEL_ARGS, "{prompt}"],
        model_args=["--model", "{model}"],
        version_cmd=["codex", "--version"],
        instructions_file="AGENTS.md",
        note="Runs Codex CLI non-interactively; Codex's own sandbox limits writes to the working directory.",
    ),
    "gemini-cli": AgentPreset(
        name="gemini-cli",
        # Without --skip-trust, Gemini CLI silently drops --yolo in an untrusted folder.
        # The prompt is the value of -p, so the model options must not come between them.
        argv=["gemini", "--yolo", "--skip-trust", "-p", "{prompt}", MODEL_ARGS],
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
    # Absolute paths: the agent runs with the working directory as its cwd, not the launch directory.
    values = {"prompt": prompt, "prompt_file": str(Path(prompt_file).resolve()),
              "workdir": str(Path(workdir).resolve()), "model": model}
    if preset is not None:
        argv: list[str] = []
        for a in preset.argv:
            if a == MODEL_ARGS:
                if model:
                    argv += [m.format(**values) for m in preset.model_args]
            else:
                argv.append(a.format(**values))
        if model and MODEL_ARGS not in preset.argv:
            argv += [m.format(**values) for m in preset.model_args]
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
                isolation: str = "unknown", ran: bool = True, status: str | None = None) -> dict[str, Any]:
    """Build a ``failbook.result/1`` record. ``status`` overrides the run status derived from the exit."""
    if status is None:
        status = run_status(exit_code, timed_out, ran)
    elif status not in RUN_STATUSES:
        raise ValueError(f"run status must be one of {', '.join(RUN_STATUSES)}, not {status!r}")
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


def _group_alive(pgid: int) -> bool:
    try:
        os.killpg(pgid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:  # macOS: a group whose members are all zombies
        return False
    return True


def _kill_group(proc: subprocess.Popen) -> bool:
    """Kill every process left in the agent's process group; True if any was still there."""
    if not hasattr(os, "killpg"):  # Windows: only the direct child can be killed
        if proc.poll() is None:
            proc.kill()
            return True
        return False
    if not _group_alive(proc.pid):
        return False
    try:
        os.killpg(proc.pid, signal.SIGKILL)
    except (ProcessLookupError, PermissionError):
        return False
    deadline = time.monotonic() + 5
    while _group_alive(proc.pid) and time.monotonic() < deadline:
        if proc.poll() is None:
            try:
                proc.wait(timeout=0.1)
            except subprocess.TimeoutExpired:
                pass
        else:
            time.sleep(0.1)
    return True


def run_agent(argv: list[str], cwd: Path, timeout: int, log_dir: Path) -> dict[str, Any]:
    """Run the agent in its own process group and make sure nothing of it is left running.

    Output goes to files rather than pipes, so a background process that inherited
    them cannot keep the run open. When the agent exits, or the time limit is hit,
    every process still in its group is killed before anything is graded.
    """
    popen_kw: dict[str, Any] = {"start_new_session": True} if hasattr(os, "killpg") else {}
    out_path, err_path = log_dir / "agent.stdout", log_dir / "agent.stderr"
    timed_out, code, notes = False, None, []
    started = time.monotonic()
    with open(out_path, "wb") as out_f, open(err_path, "wb") as err_f:
        try:
            proc = subprocess.Popen(argv, cwd=cwd, stdin=subprocess.DEVNULL, stdout=out_f, stderr=err_f, **popen_kw)
        except FileNotFoundError:
            raise SystemExit(f"error: agent command not found: {argv[0]}")
        try:
            code = proc.wait(timeout=timeout)
        except subprocess.TimeoutExpired:
            timed_out = True
        duration = time.monotonic() - started
        killed = _kill_group(proc)
        if proc.poll() is None:
            proc.wait()
    if timed_out:
        what = "its whole process group was killed" if popen_kw else "the agent process was killed"
        notes.append(f"Timed out after {timeout} s; {what} (SIGKILL) before grading.")
    elif killed:
        notes.append("The agent exited but left processes running in its process group; "
                     "they were killed before grading.")
    stdout = out_path.read_text(encoding="utf-8", errors="replace")
    stderr = err_path.read_text(encoding="utf-8", errors="replace")
    out_path.unlink()
    err_path.unlink()
    return {"stdout": stdout, "stderr": stderr, "exit_code": code, "timed_out": timed_out,
            "duration": duration, "notes": notes}


def _grader_failed(case: Case, seed: int, exc: Exception):
    from .checks import GradeReport

    first_line = (str(exc).strip().splitlines() or [""])[0][:300]
    report = GradeReport(case=case.id, case_version=case.version, seed=seed, verdict="inconclusive", rules=[],
                         engine="unknown (the grader failed)", failbook_version=__version__)
    note = (f"The grader failed ({type(exc).__name__}: {first_line}), so there is no verdict on this attempt. "
            "The work directory was kept in the run directory; re-grade it with `failbook grade`.")
    return report, note


def run_one(case: Case, *, preset: AgentPreset | None, template: str | None, agent_name: str, model: str,
            seed: int, lang: str, fix: Fix | None, timeout: int, run_dir: Path, dry_run: bool):
    from .grade import grade, prepare

    run_dir = Path(run_dir).resolve()
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
        print(f"[dry-run] cd {shlex.quote(str(work))} && {shown}")
        return None, {"command": shown}
    run = run_agent(argv, work, timeout, run_dir)
    notes = list(run["notes"])
    log = f"$ {shown}\n\n--- stdout ---\n{run['stdout']}\n--- stderr ---\n{run['stderr']}\n"
    if notes:
        log += "\n--- failbook ---\n" + "\n".join(notes) + "\n"
    (run_dir / "agent.log").write_text(log, encoding="utf-8")
    try:
        report = grade(case, work, seed)
    except CaseError:
        raise
    except Exception as exc:  # keep the attempt: the agent ran, only the grading is missing
        report, note = _grader_failed(case, seed, exc)
        notes.append(note)
    (run_dir / "grade.json").write_text(json.dumps(report.to_dict(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    info = {
        "command": shown, "exit_code": run["exit_code"], "timed_out": run["timed_out"],
        "final_message": run["stdout"].strip(), "duration": run["duration"], "inputs": inputs, "prompt": prompt,
        "outputs": {name: sha256_file(work / name) for name in case.outputs}, "notes": " ".join(notes),
    }
    return report, info


def _forbidden_runs_dir(runs_root: Path) -> Path | None:
    """The Failbook checkout (or cases directory) that ``runs_root`` lies in, if any."""
    roots = []
    root = repo_root()
    if root is not None:
        roots.append(root.resolve())
    try:
        roots.append(cases_dir().resolve())
    except CaseError:
        pass
    for r in roots:
        if runs_root == r or r in runs_root.parents:
            return r
    return None


def _enclosing_git_repo(path: Path) -> Path | None:
    for candidate in (path, *path.parents):
        if (candidate / ".git").exists():
            return candidate
    return None


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
    if args.runs_dir:
        runs_root = Path(args.runs_dir).expanduser().resolve()
        inside = _forbidden_runs_dir(runs_root)
        if inside is not None:
            print(f"error: --runs-dir {runs_root} is inside {inside}; an agent working there could read the "
                  "rules and reference solutions. Choose a directory outside it, or leave out --runs-dir "
                  "to use a new temporary directory.", file=sys.stderr)
            return 2
    else:
        runs_root = Path(tempfile.mkdtemp(prefix="failbook-runs-")).resolve()
        inside = _forbidden_runs_dir(runs_root)
        if inside is not None:
            print(f"error: the temporary directory {runs_root} is inside {inside}; pass --runs-dir with a "
                  "directory outside it.", file=sys.stderr)
            shutil.rmtree(runs_root, ignore_errors=True)
            return 2
    git_repo = _enclosing_git_repo(runs_root)
    if git_repo is not None:
        print(f"note: {runs_root} is inside the git repository {git_repo}; agents that load project instructions "
              "(AGENTS.md, CLAUDE.md, GEMINI.md) from parent directories may read that repository's files.",
              file=sys.stderr)
    fix = load_fix(args.fix) if args.fix else None
    cases = resolve_cases(args.cases or None)
    if not args.dry_run:
        from .engine import default_engine

        default_engine()  # fail before any agent runs if the grader cannot work (EngineUnavailable)
    version = agent_version(preset)
    results_dir = Path(args.results)
    if not results_dir.is_absolute() and repo_root() and not Path.cwd().joinpath(results_dir).exists():
        results_dir = repo_root() / results_dir
    print(f"Working directories and logs go to {runs_root}/", file=sys.stderr)
    summary = []
    for case in cases:
        for attempt in range(1, args.attempts + 1):
            stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
            label = "-".join(p for p in (stamp, _slug(agent_name), _slug(args.model) if args.model else "",
                                         case.id, f"a{attempt}") if p)
            runs_root.mkdir(parents=True, exist_ok=True)
            # mkdtemp adds a random suffix and fails rather than reuse a directory, so parallel runs never share one.
            run_dir = Path(tempfile.mkdtemp(prefix=label + "-", dir=runs_root))
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
                timeout_sec=args.timeout, isolation=LOCAL_ISOLATION, notes=info["notes"],
            )
            (run_dir / "result.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            path = write_result(result, results_dir)
            summary.append((case.id, attempt, report.verdict))
            print(f"{case.id} attempt {attempt}: {report.verdict.upper()} (run {result['run_status']})  -> {path}")
            if info["notes"]:
                print(f"    note: {info['notes']}")
            for r in report.rules:
                if r.status != "pass":
                    print(f"    [{r.status}] {r.id}: {r.detail}")
    if summary:
        passed = sum(v == "pass" for _, _, v in summary)
        print(f"\n{passed}/{len(summary)} attempts passed. Working directories and logs are in {runs_root}/")
    return 0
