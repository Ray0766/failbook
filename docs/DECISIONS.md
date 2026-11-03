# Decisions

Short records of the choices that shape Failbook: what was decided, the alternatives, the sources (with
versions and the date they were checked), why, how it was validated, and what is still limited. Sources
were checked on **2026-11-03**; detailed notes are in [research/2026-11-03-official-sources.md](research/2026-11-03-official-sources.md)
and [research/2026-11-03-ui-references.md](research/2026-11-03-ui-references.md). Several official doc
sites were blocked from the build environment; where that happened the same content was read from the
project's own repository, package metadata or installed source, as each note says.

---

## D1. Scope of the first release: one domain, spreadsheets

- **Decision.** 15 cases, all about editing `.xlsx` workbooks with code.
- **Who and why.** Developers and researchers who want to reproduce an agent failure, see exactly where it
  is, and check whether a fix works. Spreadsheet edits are a good first domain because the failures are
  common, subtle (the file still opens and looks fine), and *checkable deterministically* from the final file.
- **Alternatives.** Coding tasks ("tests pass but the user's action still fails"): crowded by SWE-bench-style
  benchmarks. Document layout: needs rendering comparisons, harder to check fairly. Both remain candidates.
- **Limits.** The cases are a curated set of known failure patterns, not a sample of real work
  ([DATASHEET.md](DATASHEET.md)).

## D2. Recalculate with LibreOffice; do not trust stored values

- **Problem.** A checker must know what each formula evaluates to. Files written by openpyxl store no
  results at all; files edited at the XML level can carry stale ones.
- **Decision.** Every value a rule looks at is recalculated by headless LibreOffice with
  `OOXMLRecalcMode=0` ("always recalculate", from the official schema `officecfg/.../Calc.xcs`; mode 1, the
  default, keeps stale cached values: reproduced locally). The engine and version are recorded in every result.
- **Alternatives tested in this session** (same workbook):

  | Engine | Version (release date) | XLOOKUP | TEXTJOIN | SUBTOTAL | Notes |
  |---|---|---|---|---|---|
  | pycel | 1.0b30 (2021-10-13) | no | no | yes | also returned `#VALUE!` for a valid `SUM(B2:INDEX(B:B,6))` |
  | formulas | 1.3.4 (2026-03-11) | yes | yes | no | heavier dependency tree |
  | LibreOffice | 24.2.7.2 (Ubuntu 24.04) | no (needs ≥ 24.8) | yes | yes | real spreadsheet engine |

  A wrong value from an engine bug would *fail a correct agent*, so a real application was preferred over
  a partial reimplementation.
- **Consequences.** LibreOffice is a system dependency (~1 to 4 s per recalculation, batched per file).
  Unsupported functions produce `#NAME?`; the grader reports that as **inconclusive** unless it can name a
  real cause (a formula that points at a missing sheet or name), so engine gaps never count against the agent.
  The Harbor image uses Debian trixie, whose LibreOffice 25.2 supports XLOOKUP; Ubuntu 24.04's 24.2 does not.
- **Validation.** Every reference solution also passes after LibreOffice re-saves it (`failbook verify`).
- **The audit script follows the same rule.** `audit_xlsx.py` in the skill recalculates with the same profile
  setting. When LibreOffice is unavailable, or with `--no-recalc`, it reports a `recalculation-unavailable` or
  `recalculation-skipped` warning and does not check error values at all (it does not fall back to the values
  cached in the file, which may be stale); the structural checks still run.

## D3. Deterministic rules that check the outcome, not the method

- **Decision.** Each case has 3 to 7 rules in Python. Every rule quotes the sentence of the task it checks.
  Rules have a *kind*: content, value, **live** (change an input in a copy, recalculate, check that results
  follow), **preserve** (the specific data and structures a case protects are unchanged, each rule names what it
  covers) and structure (charts, validation, formats). No LLM judge:
  nothing in these cases needs semantic judgement, and the project rules forbid using another model's
  opinion as ground truth.
- **Why live, preserve and structure rules.** A value comparison accepts numbers pasted where formulas were asked for,
  and ignores collateral damage. Measured on the 15 failure artifacts
  ([evidence/checker-ablation.md](evidence/checker-ablation.md)):

  | Grader | Failures caught | Correct solutions accepted | Empty attempts caught |
  |---|---|---|---|
  | Failbook (all rules, recalculated) | 15 / 15 | 15 / 15 | 15 / 15 |
  | values only, recalculated | 11 / 15 | 15 / 15 | 15 / 15 |
  | values only, stored values | 11 / 15 (mostly by accident) | 3 / 15 | 15 / 15 |

  This is the release's core, demonstrable claim. It is measured on reconstructed failures; it does not say
  how often agents make these mistakes.
- **Checker validation** (`failbook verify --seeds 3`, run in CI): the reference and every alternative valid
  solution (`alt_*.py`) pass, also after LibreOffice re-saves their output; the failure fails on the rules named
  in `case.yaml`; every regression fixture (`fail_*.py`, other wrong solutions found in review) is graded
  exactly `fail`; and the untouched input fails. Seeds generate different data for the same case, which
  tests that rules do not depend on one dataset and gives retests a way around memorised answers.

## D4. Run agents through Harbor (and thereby SkillEvaluator); keep a thin local runner

- **Problem.** Running third-party CLI agents fairly needs isolation, resource and time limits, repeated
  attempts and trajectory logs. Building that is out of scope.
- **Decision.** `failbook export harbor` writes standard Harbor tasks (`task.toml` schema 1.4, Dockerfile,
  `tests/test.sh` writing a flat numeric `reward.json`, `solution/solve.sh`). The grader is uploaded only
  after the agent finishes, so the agent cannot read the rules or solutions, and `test.sh` runs it with
  `python3 -I` from the tests directory, so files the agent leaves in `/workspace` cannot shadow it.
  `reward.json` always has the same keys (`reward`, `overall`, `rules_passed`, `inconclusive`,
  `grader_error`); an inconclusive or crashed grading gets reward 0 with its flag set to 1, because Harbor's
  mean (and SkillEvaluator's pass@k and lift) count it as 0. The export README explains how to exclude those
  trials. `failbook export skillevaluator`
  packages a skill fix with the cases as native Harbor evals (`grading.mode: custom_only`), so NVIDIA
  SkillEvaluator's Tier 3 can measure the skill's lift (with vs without).
- **Sources.** Harbor 0.24.0 source (installed) and PyPI (0.25.0 released 2026-10-25). SkillEvaluator 0.7.0
  (repository and PyPI, 2026-11-02), which pins `harbor==0.24.0` and supports claude-code, codex and opencode
  (not gemini-cli).
- **Validation.** All 15 exported tasks load with Harbor 0.24.0's own `Task` class; `skillevaluator tier3
  validate --harbor-contract` passes all 10 checks on an exported skill; the oracle flow (solve.sh then
  test.sh, run without Docker) gives reward 1.0.
- **Alternative considered: Inspect** (inspect-ai 0.3.287, 2026-10-31). Its Docker sandbox defaults to no
  network and routes model calls through a proxy; `inspect_swe` 0.2.72 provides Claude Code, Codex and Gemini
  agents; `inspect-harbor` 1.0.0 runs Harbor tasks inside Inspect, and its log viewer is richer. Because
  Inspect can consume Harbor tasks, exporting to Harbor keeps both paths open. Not yet tested end to end here.
- **Local runner.** `failbook run` drives an installed CLI in a fresh temporary directory outside the
  checkout (a `--runs-dir` inside it is refused, so the agent does not start next to the reference solutions).
  It is **not a sandbox**: the agent runs as your user and can read the whole file system; on a timeout the
  whole process group is killed. Presets were checked against current CLIs (Codex removed `--full-auto` in 0.147.0,
  so the preset uses `--sandbox workspace-write`; Gemini CLI needs `--skip-trust` for `--yolo` to apply;
  Claude Code refuses `--dangerously-skip-permissions` as root outside a recognised sandbox).
- **Limits.** No Docker daemon was available in the build environment, so no Harbor job was executed with a
  real agent. Agents need network access to their model API (`network_mode = "public"`).

## D5. Failure artifacts are reconstructions, labelled as such

- **Decision.** Each case ships a failure artifact produced by `failure.py`, which performs the naive steps
  (for example openpyxl `insert_rows` without rewriting totals). The case's evidence level is
  `reconstructed`; the claim shown with it is marked as written for the case. A case becomes `observed` only
  when a recorded run reproduces it ([FAILURE_TAXONOMY.md](FAILURE_TAXONOMY.md)).
- **Why.** No run records with model identifiers can be published from the environment that built this
  release, and inventing "model X said this" would be fabricated evidence.
- **Correction made on the way.** The first draft of xlsx-011 assumed "openpyxl deletes charts on save".
  `failbook verify` showed the chart survived; openpyxl 3.1.5's documentation only warns about *shapes*,
  and its source reads charts back. The case was rewritten around the real behaviour (a chart over a fixed
  range does not grow with new rows).

## D6. Results: append-only records, counts not rates, no ranking

- **Decision.** One JSON file per attempt in `results/<case>/`, never edited or overwritten (a review adds a
  field or a new file). `failbook run` fills the prompt and input hashes for every seed; `failbook grade
  --record` takes the prompt from the directory's `TASK.md` or `--prompt-file` and records timeouts and agent
  errors with `--run-status`. Timeouts and agent errors are reported apart from task failures. Each record has the agent and version, model, fix, harness, seed, run status
  (completed / timeout / agent-error), a mechanical classification, input/output/prompt SHA-256, environment,
  budget (`null` when unknown) and the final message. `RESULTS.md` shows *passed / attempts* per case and
  configuration and deliberately has no cross-case total or ranking.
- **Why.** A curated collection of failure patterns does not support statements about overall reliability.

## D7. A fix is something you can retest: the spreadsheet-verify skill and its audit script

- **Decision.** Fixes live in `fixes/`. The first one is an Agent Skill: a checklist of the pitfalls plus a
  task-agnostic audit script (`scripts/audit_xlsx.py`, generated from `src/failbook/audit.py`) that compares
  the edited file with the original.
- **Validation** ([evidence/fix-self-check.md](evidence/fix-self-check.md)): the script flags 13 of the 15
  failure artifacts (9 with an error-level finding) and reports no error on any of the 30 correct solutions
  (one warning on a valid "static values" alternative, a symptom the original file already had). It misses
  xlsx-002 and xlsx-008, which leave no generic symptom.
- **Corrected after the review (D9).** The audit raised errors on ordinary correct edits (stacked summary
  rows below a total, a renamed header or row label, a symptom the original already had, `#NAME?` from a
  function LibreOffice lacks) and crashed when LibreOffice could not start. Those are now warnings or a
  stated `recalculation-unavailable` finding; real breakage (a shifted column, deleted rows, a stale sheet
  name) is still an error. Each case is a test in `tests/test_audit_more.py`.
- **Limits.** The checks were written with these failures in view, so this is coverage, not generalisation.
  **Whether agents do better with the skill is a hypothesis** until runs with and without it are recorded.

## D8. Site: a "research archive", not a dashboard

- **Directions compared.** (A) document-style case dossiers with a dense catalog table; (B) a debugging
  workbench with a rule list and an evidence pane.
- **References actually viewed** (2026-11-03): the Inspect log viewer (inspect-ai 0.3.287, run locally), the
  Playwright HTML report (1.63.0), a Lighthouse 13.6.0 HTML report, and jsr.io's score and versions pages.
  The pages named in the quality brief (inspect.aisi.org.uk, swebench.com, programbench.com) and the AI
  Incident Database and caniuse were blocked by the network policy and were **not** seen.
- **Choice: A.** Visitors arrive from a link to understand one failure; a linear page (task, expected,
  observed, evidence, reproduce, retests) matches that order, works without JavaScript, prints, and reflows
  to 320 px. Both workbench references broke at 390 px (Inspect squeezed its explanation column to one
  character per line). Borrowed: Lighthouse's failures-first list with passed items collapsed and counted,
  and its caveat next to the number; Inspect's evidence string next to each score and prev/next navigation;
  Playwright's searchable filter chips with counts and labelled expected/observed values. Avoided: one icon
  for "completed" and "passed" (Inspect), means over three samples, a headline score that contradicts a
  visible failure (JSR), card walls and relative dates.
- **Visual system.** System fonts with CJK fallbacks, 13 to 30 px scale, colour only for verdicts, links and
  focus; verdicts are glyph + word (WCAG 2.2 1.4.1), 2 px focus ring, no page-level horizontal scroll at
  320 px (checked with Playwright on five pages), `prefers-reduced-motion` respected. Not a full WCAG audit.

## D9. An independent adversarial review before release

- **What.** Before merging, eight reviewer agents with separate lenses (checker fairness for cases 001-008 and
  009-015, core bugs, accuracy of public claims) tried to break the release. Each finding needed reproduction
  commands; a second, skeptical agent per lens tried to refute it. 52 findings were reported and 51 survived.
- **Examples.** Graders that let partly hard-coded answers pass because a live rule perturbed one row; a stale
  defined name graded inconclusive instead of fail; a chart over a named range graded fail; Harbor's
  `test.sh` importable from the agent's workspace; `failbook run` defaulting to a directory inside the
  checkout; "observed" texts that did not match the committed artifacts; a headline sentence contradicted by
  its own table.
- **Handling.** Each was fixed in the file group that owns it. Wrong solutions the reviewers wrote became
  `fail_*.py` regression fixtures and valid ones `alt_*.py`, so `failbook verify` keeps checking them.
- **Limits.** The reviewers are models, not ground truth; their findings were accepted only with a reproduced
  failure. After the fixes, each finding's reproduction was run again (wrong solutions now fail, valid ones
  pass or are inconclusive with a stated reason; logs were kept outside the repository). A fresh,
  independent review of the fixed code has not been run.

## D10. Grading policies made explicit by the review

The review showed that several graders encoded a policy nobody had written down. These are now decisions,
each backed by an `alt_*.py` (accepted) or `fail_*.py` (rejected) fixture that `failbook verify` re-runs:

| Policy | Cases | Fixture |
|---|---|---|
| Liveness checks change many inputs at once (every row in most cases; a different row per quarter in xlsx-010; the rate in xlsx-003) and check every dependent figure, so a partly typed answer does not pass | 001, 002, 003, 005, 006, 007, 008, 009, 010, 014 | e.g. `fail_partial_values.py`, `fail_typed_new_amounts.py` |
| Money may be rounded to whole cents per line (at most 1 cent from exact) if the totals agree with those lines; the tasks do not forbid it and it is common accounting practice | 003, 005 | `alt_round_cents.py` |
| A ratio the task says not to round must not be rounded (the task text now says so in both languages) | 002 | `fail_rounded_margin.py` |
| A valid form the checker cannot read (a chart series over an `OFFSET` name, a conditional-format type the evaluator does not implement) is **inconclusive** with the exact reason, never fail | 011, 013 | `alt_named_range.py`, `alt_begins_with.py` |
| What the user sees is checked, not only that something exists: chart series start at the first data row, carry the column's legend name and month labels; conditional formats are evaluated per cell (priority, `stopIfTrue`, colour resolved to a hue) | 011, 013 | `fail_header_in_values.py`, `fail_green_highlight.py` |

- **Cross-check of the conditional-format evaluator.** For 10 solutions of xlsx-013 (seed 0) the checker's
  per-cell "red or not" agreed with the background colour LibreOffice paints when it exports HTML on 170 of
  170 cells. This is agreement with one engine, not with Excel.
- **Case versions stay at 1.** Every rule and task change above was made before the first release, and no
  result record references any case version, so there is nothing that a version bump would keep apart. From
  the first release on, a change to a task or a rule bumps the version ([DATASHEET.md](DATASHEET.md)).
- **Limits.** Each policy is a judgment about what a careful person would accept; a contributor can contest
  one through the "checker dispute" issue form.
