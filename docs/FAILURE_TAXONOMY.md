# Failure taxonomy and evidence levels

## Who failed: `classification.category` (per case) and `classification.reviewed` (per run)

| Category | Meaning | Example |
|---|---|---|
| `model` | The model or agent made a wrong decision: wrong approach, skipped verification, claimed success without checking. | Writing numbers as text; summing subtotals; not extending a chart range. |
| `tool` | A tool or library behaved incorrectly or contrary to its documentation. | A library that corrupts a file it claims to support. |
| `environment` | The run itself broke: missing dependency, timeout, crash, network or quota error. | The agent CLI exits with an error before editing anything. |
| `evaluator` | The checker was wrong or could not decide. | An engine that does not implement a function the agent used. |
| `undetermined` | Not yet analysed. | |

Documented tool behaviour that the agent did not account for (openpyxl's `insert_rows` not rewriting
formulas, pandas' `to_excel` writing a new file) is a **model** failure with the tool behaviour recorded in
`classification.tool_behaviour`; it is not a tool bug.

## How strong the evidence is: `classification.evidence`

| Level | Meaning |
|---|---|
| `reconstructed` | The failure artifact was produced by `failure.py`, a scripted version of the naive approach. It shows the failure is possible and checkable, not that a particular agent makes it. |
| `observed` | At least one recorded run in `results/` produced a failing output for this case, and a reviewer confirmed the failure matches the case. The case links the run. |

## What happened in one run: result fields

- `run_status`: `completed`, `timeout`, `agent-error` (non-zero exit), or `not-run`.
- `verdict` (from the grader): `pass`, `fail`, or `inconclusive` (never counted as a failure).
- `classification.auto`, assigned mechanically: `none` (pass), `task-failure` (completed but failed the
  rules), `agent-or-environment-error` (timeout or agent error), `evaluator-inconclusive`.
- `classification.reviewed`: filled by a human reviewer with one of the categories above, plus a note.
  Machine and human labels are kept apart.
