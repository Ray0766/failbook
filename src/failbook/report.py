"""Aggregating result files into RESULTS.md."""

from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path
from typing import Any

from .case import all_cases
from .runner import RUN_STATUSES

REQUIRED = ["schema", "case", "case_version", "seed", "agent", "model", "harness", "started_at", "verdict", "rules"]
VERDICTS = {"pass", "fail", "inconclusive"}
#: Agent runs that did not complete: the agent or its environment failed. Never counted as task failures.
RUN_PROBLEMS = tuple(s for s in RUN_STATUSES if s != "completed")


def validate_result(data: Any) -> list[str]:
    problems = []
    if not isinstance(data, dict):
        return ["not a JSON object"]
    for key in REQUIRED:
        if key not in data:
            problems.append(f"missing key {key!r}")
    if data.get("schema") != "failbook.result/1":
        problems.append(f"unknown schema {data.get('schema')!r}")
    if data.get("verdict") not in VERDICTS:
        problems.append(f"verdict must be one of {sorted(VERDICTS)}")
    if data.get("run_status", "completed") not in RUN_STATUSES:
        problems.append(f"run_status must be one of {list(RUN_STATUSES)}")
    agent = data.get("agent")
    if not isinstance(agent, dict) or not agent.get("name"):
        problems.append("agent.name is required")
    for r in data.get("rules") or []:
        if not isinstance(r, dict) or r.get("status") not in VERDICTS:
            problems.append("every rule needs a status of pass, fail or inconclusive")
            break
    return problems


def load_results(results_dir: Path) -> tuple[list[dict], list[str]]:
    results, problems = [], []
    if not Path(results_dir).is_dir():
        return results, problems
    for path in sorted(Path(results_dir).rglob("*.json")):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except ValueError as exc:
            problems.append(f"{path}: invalid JSON ({exc})")
            continue
        issues = validate_result(data)
        if issues:
            problems.append(f"{path}: " + "; ".join(issues))
            continue
        data["_path"] = str(path)
        results.append(data)
    return results, problems


def config_label(r: dict) -> str:
    label = r["agent"]["name"]
    if r.get("model"):
        label += f" · {r['model']}"
    if r.get("fix"):
        label += f" + {r['fix']}"
    return label


def outcome(r: dict) -> str:
    """What one attempt counts as: its run status if the agent run did not complete, else the verdict."""
    status = r.get("run_status") or "completed"
    return status if status in RUN_PROBLEMS else r["verdict"]


def summarize(results: list[dict]) -> dict[str, Any]:
    """Count outcomes per configuration and case.

    ``table`` holds the current case versions; ``older`` holds every other version, keyed by
    (case, case_version), so that results graded against an old task or old rules stay visible
    without being mixed in.
    """
    versions = {c.id: c.version for c in all_cases()}
    table: dict[str, dict[str, dict[str, int]]] = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    older: dict[tuple[str, Any], dict[str, dict[str, int]]] = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    stale = 0
    for r in results:
        if versions.get(r["case"]) == r["case_version"]:
            cell = table[config_label(r)][r["case"]]
        else:
            stale += 1
            cell = older[(r["case"], r["case_version"])][config_label(r)]
        cell["attempts"] += 1
        cell[outcome(r)] += 1
    return {"table": table, "older": older, "stale": stale, "total": len(results), "versions": versions}


def _cell(c: dict[str, int] | None) -> str:
    if not c:
        return "·"
    passed, decided = c.get("pass", 0), c.get("pass", 0) + c.get("fail", 0)
    uncounted = [f"{c[k]} {k}" for k in ("inconclusive", *RUN_PROBLEMS) if c.get(k)]
    if not decided:
        return "❔ " + ", ".join(uncounted)
    mark = "✅" if passed == decided else ("❌" if not passed else "⚠️")
    return f"{mark} {passed}/{decided}" + (f" ({', '.join(uncounted)})" if uncounted else "")


def _older_section(s: dict[str, Any], cases: list) -> list[str]:
    if not s["older"]:
        return []
    by_id = {c.id: c for c in cases}
    configs = sorted({cfg for row in s["older"].values() for cfg in row})
    lines = [
        "## Results for older case versions",
        "",
        "These attempts were graded against a version of the case that is not the current one: its task or "
        "its rules have changed since. They are kept, but not counted in the table above.",
        "",
        "| Case | Graded version | Current version | " + " | ".join(configs) + " |",
        "|---|---|---|" + "---|" * len(configs),
    ]
    for (cid, version) in sorted(s["older"], key=lambda k: (k[0], str(k[1]))):
        case = by_id.get(cid)
        name = f"[{cid}](cases/{case.dir.name}/) {case.title}" if case else f"{cid} (not in this checkout)"
        current = str(case.version) if case else "–"
        row = [name, str(version), current] + [_cell(s["older"][(cid, version)].get(cfg)) for cfg in configs]
        lines.append("| " + " | ".join(row) + " |")
    lines.append("")
    return lines


def render_markdown(results: list[dict]) -> str:
    cases = all_cases()
    s = summarize(results)
    lines = [
        "# Retest results",
        "",
        "<!-- generated by `failbook report`; do not edit by hand -->",
        "",
        "Each cell is *passed / graded attempts* for the current version of the case: attempts whose agent run "
        "completed and that the grader could decide. ✅ all passed, ⚠️ some passed, ❌ none passed, ❔ no attempt "
        "could be graded. Attempts that were not graded are listed by reason and never counted as failures: "
        "`inconclusive` (the grader could not decide), `timeout`, `agent-error` and `not-run` (the agent run did "
        "not complete: an agent or environment problem, see `run_status`). How results are produced and "
        "submitted: [results/README.md](results/README.md).",
        "",
        "There is deliberately no overall score or ranking. The cases are a curated set of known failure "
        "patterns, not a sample of real work, so a pass count across cases says nothing about how reliable an "
        "agent is in general (see [docs/DATASHEET.md](docs/DATASHEET.md)).",
        "",
    ]
    if not results:
        lines += [
            "**No retest results have been submitted yet.** The failure artifacts shipped with each case are "
            "reconstructions (see each case's provenance), not measurements of any particular model.",
            "",
            "Be the first: run `failbook run --agent <preset> --model <name>`, or record a run made with any "
            "other tool with `failbook grade --record` (see [results/README.md](results/README.md)), and open a "
            "pull request with the files written to `results/`.",
            "",
        ]
        return "\n".join(lines)
    if s["table"]:
        configs = sorted(s["table"])
        lines.append("| Case | " + " | ".join(configs) + " |")
        lines.append("|---|" + "---|" * len(configs))
        for c in cases:
            row = [f"[{c.id}](cases/{c.dir.name}/) {c.title}"]
            row += [_cell(s["table"][cfg].get(c.id)) for cfg in configs]
            lines.append("| " + " | ".join(row) + " |")
        lines.append("")
    else:
        lines += ["No results for the current case versions yet.", ""]
    lines += _older_section(s, cases)
    current = s["total"] - s["stale"]
    lines += [f"{s['total']} result file(s): {current} for the current case versions, "
              f"{s['stale']} for older ones.", ""]
    return "\n".join(lines)


def write_report(results_dir: Path, out: Path | None) -> int:
    """Write RESULTS.md (or print it). Returns 1 if any result file is invalid; those are left out."""
    import sys

    results, problems = load_results(results_dir)
    for p in problems:
        print(f"invalid result: {p}", file=sys.stderr)
    text = render_markdown(results)
    if out:
        out.write_text(text, encoding="utf-8")
        print(f"wrote {out} from {len(results)} result file(s)")
    else:
        print(text)
    if problems:
        print(f"error: {len(problems)} invalid result file(s) were left out", file=sys.stderr)
    return 1 if problems else 0
