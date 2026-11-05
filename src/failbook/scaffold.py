"""`failbook new`: scaffold a case directory with every file a case needs."""

from __future__ import annotations

import re
from pathlib import Path

from .case import CaseError, cases_dir

CASE_YAML = """id: {id}
slug: {slug}
version: 1
status: draft   # draft | checker-verified | observed | retired
domain: spreadsheet
difficulty: easy
tags: []
title: Describe the observable failure in one sentence
title_zh: 用一句话描述可观察到的失败
files:
  inputs: [data.xlsx]
  outputs: [data.xlsx]
failure:
  expected: >
    What a correct result looks like.
  expected_zh: >
    正确结果是什么样子。
  observed: >
    What the failure artifact contains (produced by failure.py).
  observed_zh: >
    失败产物里实际是什么。
  claim: >
    The kind of success message that accompanies this failure (illustrative).
  claim_zh: 与该失败相伴的示例声明。
  explanation: >
    Why it happens and why it is easy to miss.
  explanation_zh: >
    为什么会发生、为什么容易被忽略。
  failing_rules: [result-correct]
classification:
  category: undetermined   # model | tool | environment | evaluator | undetermined
  evidence: reconstructed   # reconstructed | observed
  tool_behaviour: >
    Which tool behaviour is involved, if any.
limitations:
  - en: What the rules do not check.
    zh: 规则没有检查什么。
inputs:
  source: synthetic
  license: MIT
provenance:
  kind: reconstructed
  note: >
    The failure artifact is produced by failure.py. It is not attributed to any specific model run.
"""

CASE_PY = '''import openpyxl

from failbook import kit
from failbook import xlsx as X
from failbook.checks import require, rule

SHEET = "Data"


def params(seed):
    r = kit.rng(seed, "{id}")
    return {{"values": [r.randint(1, 100) for _ in range(5)]}}


def build(meta, outdir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.append(["Value"])
    kit.style_header(ws, 1, 1)
    for v in meta["values"]:
        ws.append([v])
    wb.save(outdir / "data.xlsx")


@rule(
    "result-correct",
    "Describe what this rule checks",
    kind="value",
    title_zh="这条规则检查什么",
    requirement="Quote the sentence of the task this rule verifies.",
)
def result_correct(ctx):
    values = ctx.values()
    got = ctx.num(values, X.make_ref(SHEET, "B1"), "the result")
    require(X.close(got, sum(ctx.meta["values"])), f"the result is {{got}}, expected {{sum(ctx.meta['values'])}}")
'''

SOLUTION = '''"""{doc}"""

import openpyxl


def solve(workdir, meta):
    path = workdir / "data.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Data"]
    ws["B1"] = {value}
    wb.save(path)
'''


def new_case(case_id: str, slug: str) -> Path:
    if not re.fullmatch(r"[a-z]+-\d{3}", case_id):
        raise CaseError("case id must look like xlsx-016")
    if not re.fullmatch(r"[a-z0-9]+(-[a-z0-9]+)*", slug):
        raise CaseError("slug must be lowercase words joined by hyphens")
    root = cases_dir()
    if any(p.name.startswith(case_id + "-") for p in root.iterdir()):
        raise CaseError(f"{case_id} already exists")
    d = root / f"{case_id}-{slug}"
    d.mkdir()
    (d / "case.yaml").write_text(CASE_YAML.format(id=case_id, slug=slug), encoding="utf-8")
    (d / "case.py").write_text(CASE_PY.format(id=case_id), encoding="utf-8")
    (d / "task.md").write_text("`data.xlsx` has numbers in column A (sheet **Data**). Put their total in B1 as a formula.\n", encoding="utf-8")
    (d / "task.zh.md").write_text("`data.xlsx`（工作表 **Data**）的 A 列是数字。请在 B1 用公式写出它们的合计。\n", encoding="utf-8")
    (d / "reference.py").write_text(SOLUTION.format(doc="A correct solution.", value='"=SUM(A2:A6)"'), encoding="utf-8")
    (d / "failure.py").write_text(SOLUTION.format(doc="Reproduces the failure.", value='"=SUM(A2:A5)"'), encoding="utf-8")
    return d
