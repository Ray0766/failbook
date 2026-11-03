"""Command line interface: ``failbook <command>``."""

from __future__ import annotations

import argparse
import json
import shutil
import sys
import tempfile
from pathlib import Path

from . import __version__
from .case import Case, CaseError, resolve_cases


def _print_report(case: Case, report, *, lang: str = "en") -> None:
    icon = {"pass": "PASS", "fail": "FAIL", "inconclusive": "????"}
    title = case.title_zh if lang == "zh" else case.title
    print(f"{case.id}  {title}")
    print(f"verdict: {report.verdict.upper()}   (seed {report.seed}, engine: {report.engine})")
    for r in report.rules:
        print(f"  [{icon[r.status]}] {r.id}: {r.title}")
        if r.detail:
            print(f"         {r.detail}")


def cmd_list(args) -> int:
    cases = resolve_cases(None)
    if args.json:
        print(json.dumps([{"id": c.id, "slug": c.slug, "title": c.title, "title_zh": c.title_zh,
                           "tags": c.spec.get("tags", []), "difficulty": c.spec.get("difficulty"),
                           "status": c.spec.get("status")} for c in cases], ensure_ascii=False, indent=2))
        return 0
    for c in cases:
        title = c.title_zh if args.lang == "zh" else c.title
        print(f"{c.id}  {c.spec.get('difficulty', ''):<6}  {title}")
    return 0


def cmd_show(args) -> int:
    for case in resolve_cases(args.cases):
        meta = case.params(args.seed)
        lang = args.lang
        print(f"# {case.id}: {case.title_zh if lang == 'zh' else case.title}\n")
        print("## Task\n")
        print(case.prompt(meta, lang))
        print("## Expected\n")
        print(case.text("failure.expected", lang) + "\n")
        print("## Observed in the failure artifact (reconstructed)\n")
        print(case.text("failure.observed", lang) + "\n")
        print("## Rules\n")
        for r in case.rules:
            print(f"- {r.id} [{r.kind}]: {r.title_zh if lang == 'zh' and r.title_zh else r.title}")
        print()
    return 0


def cmd_prepare(args) -> int:
    from .grade import prepare

    case = resolve_cases([args.case])[0]
    out = Path(args.dir)
    if out.exists() and any(out.iterdir()) and not args.force:
        print(f"error: {out} is not empty (use --force)", file=sys.stderr)
        return 2
    prepare(case, out, args.seed, lang=args.lang)
    print(f"Prepared {case.id} (seed {args.seed}) in {out}")
    print(f"Give the agent the task in {out / 'TASK.md'}, then run:")
    print(f"  failbook grade {case.id} {out}" + (f" --seed {args.seed}" if args.seed else ""))
    return 0


def cmd_grade(args) -> int:
    from .grade import grade

    case = resolve_cases([args.case])[0]
    report = grade(case, Path(args.dir), args.seed)
    if args.json:
        Path(args.json).write_text(json.dumps(report.to_dict(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    if args.harbor_reward:
        from .harbor import write_reward

        write_reward(report, Path(args.harbor_reward))
    _print_report(case, report, lang=args.lang)
    if args.record:
        from .runner import _now, make_result, sha256_file, write_result

        if not args.agent:
            print("error: --record needs --agent (and ideally --model)", file=sys.stderr)
            return 2
        result = make_result(
            case, report, seed=args.seed, lang=args.lang, agent=args.agent, agent_ver=args.agent_version,
            command="", model=args.model, fix=args.fix, harness=args.harness, attempt=1, attempts=1,
            started_at=_now(), duration=None, exit_code=None, timed_out=False,
            final_message=Path(args.final_message).read_text(encoding="utf-8") if args.final_message else "",
            submitter=args.submitter, notes=args.notes,
            outputs={n: sha256_file(Path(args.dir) / n) for n in case.outputs},
            isolation=args.isolation,
        )
        print(f"recorded {write_result(result, Path(args.record))}")
    return {"pass": 0, "fail": 1}.get(report.verdict, 3)


def cmd_solve(args) -> int:
    case = resolve_cases([args.case])[0]
    meta = case.params(args.seed)
    case.solve(args.kind, Path(args.dir), meta)
    print(f"Applied the {args.kind} solution of {case.id} to {args.dir}")
    return 0


def cmd_verify(args) -> int:
    from .grade import verify

    seeds = list(range(args.seeds)) if args.seeds else [0]
    bad = 0
    for case in resolve_cases(args.cases):
        try:
            outcomes = verify(case, seeds)
        except Exception as exc:
            bad += 1
            print(f"ERROR {case.id}: {type(exc).__name__}: {exc}")
            continue
        for o in outcomes:
            if o.ok:
                failed = [r.id for r in o.failure.rules if r.status == "fail"]
                alts = f" + {len(o.alternatives)} alternative(s)" if o.alternatives else ""
                print(f"ok    {case.id} seed {o.seed}: reference{alts} pass (also after re-save), "
                      f"failure caught by {', '.join(failed)}")
            else:
                bad += 1
                print(f"FAIL  {case.id} seed {o.seed}:")
                for p in o.problems:
                    print(f"        - {p}")
    if args.check_committed:
        from .build import check_committed

        for case in resolve_cases(args.cases):
            problems = check_committed(case)
            for p in problems:
                bad += 1
                print(f"STALE {case.id}: {p}")
    return 1 if bad else 0


def cmd_build(args) -> int:
    from .build import build_case

    for case in resolve_cases(args.cases):
        build_case(case)
        print(f"built {case.id}")
    if not args.no_docs:
        from .docs import write_all_docs

        write_all_docs()
        print("docs updated")
    return 0


def cmd_docs(args) -> int:
    from .docs import write_all_docs

    changed = write_all_docs(check=args.check)
    if args.check and changed:
        print("out of date: " + ", ".join(str(p) for p in changed))
        print("run `failbook docs` and commit the result")
        return 1
    print("docs up to date" if args.check else "docs written")
    return 0


def cmd_run(args) -> int:
    from .runner import run_cases

    return run_cases(args)


def cmd_export(args) -> int:
    from .harbor import export_harbor, export_skillevaluator

    cases = resolve_cases(args.cases)
    out = Path(args.out)
    if args.format == "harbor":
        export_harbor(cases, out, seed=args.seed, lang=args.lang)
        print(f"Exported {len(cases)} Harbor task(s) to {out}")
    else:
        if not args.fix:
            print("error: --fix is required for the skillevaluator format", file=sys.stderr)
            return 2
        path = export_skillevaluator(cases, out, fix=args.fix, seed=args.seed, lang=args.lang)
        print(f"Exported a SkillEvaluator skill with {len(cases)} native Harbor task(s) to {path}")
    return 0


def cmd_report(args) -> int:
    from .report import write_report

    write_report(Path(args.results), Path(args.out) if args.out else None)
    return 0


def cmd_site(args) -> int:
    from .site import build_site

    path = build_site(Path(args.out), results_dir=Path(args.results))
    print(f"Site written to {path}")
    return 0


def cmd_ablation(args) -> int:
    from .ablation import write_ablation

    path = write_ablation(resolve_cases(args.cases), Path(args.out))
    print(path.read_text(encoding="utf-8").split("## Per case")[0])
    print(f"written to {path}")
    return 0


def cmd_audit(args) -> int:
    from .audit import main as audit_main

    argv = [args.after] + (["--before", args.before] if args.before else []) + (["--json"] if args.json else [])
    return audit_main(argv)


def cmd_audit_eval(args) -> int:
    from .ablation import write_audit_eval

    path = write_audit_eval(resolve_cases(args.cases), Path(args.out))
    print(path.read_text(encoding="utf-8").split("## Per case")[0])
    print(f"written to {path}")
    return 0


def cmd_new(args) -> int:
    from .scaffold import new_case

    path = new_case(args.id, args.slug)
    print(f"Created {path}. Edit case.yaml, task.md, case.py, reference.py and failure.py, then run:")
    print(f"  failbook build {args.id} && failbook verify {args.id} --seeds 3")
    return 0


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="failbook", description="Reproducible AI failure cases with fair, re-runnable checks.")
    p.add_argument("--version", action="version", version=f"failbook {__version__}")
    sub = p.add_subparsers(dest="command", required=True)

    def lang_opt(sp):
        sp.add_argument("--lang", choices=["en", "zh"], default="en", help="task / output language")

    sp = sub.add_parser("list", help="list cases")
    sp.add_argument("--json", action="store_true")
    lang_opt(sp)
    sp.set_defaults(func=cmd_list)

    sp = sub.add_parser("show", help="print a case's task, failure and rules")
    sp.add_argument("cases", nargs="*")
    sp.add_argument("--seed", type=int, default=0)
    lang_opt(sp)
    sp.set_defaults(func=cmd_show)

    sp = sub.add_parser("prepare", help="create a working directory for a manual attempt")
    sp.add_argument("case")
    sp.add_argument("dir")
    sp.add_argument("--seed", type=int, default=0, help="variant seed (0 = canonical)")
    sp.add_argument("--force", action="store_true")
    lang_opt(sp)
    sp.set_defaults(func=cmd_prepare)

    sp = sub.add_parser("grade", help="grade the output files in a directory")
    sp.add_argument("case")
    sp.add_argument("dir")
    sp.add_argument("--seed", type=int, default=0)
    sp.add_argument("--json", metavar="PATH", help="also write the report as JSON")
    sp.add_argument("--harbor-reward", metavar="DIR", help="write Harbor reward.txt/reward.json into DIR")
    sp.add_argument("--record", metavar="RESULTS_DIR", help="also write a result file (for manual or external runs)")
    sp.add_argument("--agent", default="", help="(with --record) the agent or app used, e.g. chatgpt-web, cursor")
    sp.add_argument("--agent-version", default="")
    sp.add_argument("--model", default="", help="(with --record) the model name as the tool reports it")
    sp.add_argument("--fix", default=None, help="(with --record) the fix applied, if any")
    sp.add_argument("--harness", default="manual", help="(with --record) how the run was done: manual, harbor, ...")
    sp.add_argument("--final-message", metavar="FILE", help="(with --record) file with the agent's final reply")
    sp.add_argument("--isolation", default="unknown", help="(with --record) how the agent was isolated, e.g. docker")
    sp.add_argument("--submitter", default="")
    sp.add_argument("--notes", default="")
    lang_opt(sp)
    sp.set_defaults(func=cmd_grade)

    sp = sub.add_parser("solve", help="apply the reference or failure solution to a directory")
    sp.add_argument("case")
    sp.add_argument("dir")
    sp.add_argument("--kind", choices=["reference", "failure"], default="reference")
    sp.add_argument("--seed", type=int, default=0)
    sp.set_defaults(func=cmd_solve)

    sp = sub.add_parser("verify", help="prove each checker accepts the reference and catches the failure")
    sp.add_argument("cases", nargs="*")
    sp.add_argument("--seeds", type=int, default=1, help="number of variant seeds to test (0..N-1)")
    sp.add_argument("--check-committed", action="store_true", help="also check committed inputs/artifacts are current")
    sp.set_defaults(func=cmd_verify)

    sp = sub.add_parser("build", help="regenerate committed inputs, failure artifacts and docs")
    sp.add_argument("cases", nargs="*")
    sp.add_argument("--no-docs", action="store_true")
    sp.set_defaults(func=cmd_build)

    sp = sub.add_parser("docs", help="regenerate case READMEs and the case index")
    sp.add_argument("--check", action="store_true", help="exit 1 if any generated doc is out of date")
    sp.set_defaults(func=cmd_docs)

    sp = sub.add_parser("run", help="run an agent on cases and record results")
    sp.add_argument("cases", nargs="*")
    sp.add_argument("--agent", required=True, help="preset name (claude-code, codex, gemini-cli) or 'custom'")
    sp.add_argument("--cmd", help="command template for --agent custom; placeholders {prompt}, {prompt_file}, {workdir}, {model}")
    sp.add_argument("--model", default="", help="model name passed to the agent and recorded in the result")
    sp.add_argument("--seed", type=int, default=0)
    sp.add_argument("--attempts", type=int, default=1, help="attempts per case (all are recorded)")
    sp.add_argument("--fix", help="apply a fix from fixes/ (skill or prompt) before running")
    sp.add_argument("--timeout", type=int, default=900, help="seconds per attempt")
    sp.add_argument("--results", default="results", help="directory to write result files into")
    sp.add_argument("--runs-dir", default="runs", help="where to keep working directories and logs")
    sp.add_argument("--submitter", default="", help="your GitHub handle, recorded in the result")
    sp.add_argument("--dry-run", action="store_true", help="prepare and print the command without running it")
    lang_opt(sp)
    sp.set_defaults(func=cmd_run)

    sp = sub.add_parser("export", help="export cases to Harbor tasks or a SkillEvaluator skill")
    sp.add_argument("format", choices=["harbor", "skillevaluator"])
    sp.add_argument("cases", nargs="*")
    sp.add_argument("--out", required=True)
    sp.add_argument("--seed", type=int, default=0)
    sp.add_argument("--fix", help="(skillevaluator) the skill fix to evaluate")
    lang_opt(sp)
    sp.set_defaults(func=cmd_export)

    sp = sub.add_parser("report", help="summarise results/ into RESULTS.md")
    sp.add_argument("--results", default="results")
    sp.add_argument("--out", default="RESULTS.md")
    sp.set_defaults(func=cmd_report)

    sp = sub.add_parser("site", help="build the static case gallery")
    sp.add_argument("--out", default="_site")
    sp.add_argument("--results", default="results")
    sp.set_defaults(func=cmd_site)

    sp = sub.add_parser("ablation", help="compare the full grader with weaker baseline graders")
    sp.add_argument("cases", nargs="*")
    sp.add_argument("--out", default="docs/evidence")
    sp.set_defaults(func=cmd_ablation)

    sp = sub.add_parser("audit", help="check an edited workbook for silent spreadsheet mistakes (case-agnostic)")
    sp.add_argument("after")
    sp.add_argument("--before", help="the original workbook, to detect collateral damage")
    sp.add_argument("--json", action="store_true")
    sp.set_defaults(func=cmd_audit)

    sp = sub.add_parser("audit-eval", help="measure the audit on failure artifacts and correct solutions")
    sp.add_argument("cases", nargs="*")
    sp.add_argument("--out", default="docs/evidence")
    sp.set_defaults(func=cmd_audit_eval)

    sp = sub.add_parser("new", help="scaffold a new case")
    sp.add_argument("id", help="e.g. xlsx-016")
    sp.add_argument("slug", help="e.g. rounding-penny-gap")
    sp.set_defaults(func=cmd_new)

    args = p.parse_args(argv)
    try:
        return args.func(args)
    except CaseError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    except BrokenPipeError:  # output piped into `head` and friends
        sys.stderr.close()
        return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
