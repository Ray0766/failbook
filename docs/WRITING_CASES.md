# Writing a case and its checks

1. `failbook new xlsx-016 short-slug` scaffolds `cases/xlsx-016-short-slug/`.
2. `case.py`: `params(seed)` returns the data for one variant (use `kit.rng(seed, "xlsx-016")`), `build(meta, outdir)`
   writes the input workbook, optional `scenarios(ctx)` declares perturbations, and `@rule` functions check the result.
3. `task.md` / `task.zh.md`: the task. `{{name}}` placeholders are filled from `params()`.
4. `reference.py` (a correct solution), `failure.py` (the failure, reproduced), optional `alt_*.py` (other valid
   solutions, which must pass) and `fail_*.py` (other wrong solutions, which must be graded exactly `fail`).
5. `case.yaml`: titles, expected / observed / explanation, the illustrative claim, `failure.failing_rules`,
   classification, limitations, provenance.
6. `failbook build xlsx-016 && failbook verify xlsx-016 --seeds 3`.

## Rules for rules

- **Trace every rule to the task.** `requirement=` quotes the sentence it checks. If the task does not ask for
  formulas, do not require formulas (xlsx-015 accepts correct static values; `alt_static_values.py` proves it).
  The footer every task carries ("do not change anything the task does not ask you to change") is what
  `preserve` rules rely on.
- **Check outcomes, not methods.** Compare recalculated values, never formula text. Test liveness by
  perturbing inputs (`scenarios`) rather than by looking for `=` signs.
- **Compute expectations independently.** Expected values come from `params()` in Python, not from the engine.
- **Fail with evidence.** Messages name the cell and both numbers: `Total Qty is 198, expected 269 (difference 71)`.
- **Never fail the agent for the checker's limits.** Use `ctx.num()` / `ctx.value()`, which turn engine gaps
  into `Inconclusive`; let unexpected exceptions surface as inconclusive "checker error".
- **Add an alternative solution** whenever there is a reasonable different way to do the task, and a `fail_*.py`
  fixture for every plausible wrong answer a rule is meant to catch. A live rule that perturbs one row lets
  an answer that is a formula only on that row pass; perturb every row, or say so under limitations.
- **Describe the artifact, not a story.** `observed` says what the committed failure file contains (cells,
  formulas, values), checked against the file; the illustrative `claim` is labelled as such.
- **Prove it.** `verify` must pass on at least three seeds, and the failure must fail on exactly the rules
  that explain it (`failure.failing_rules`).
