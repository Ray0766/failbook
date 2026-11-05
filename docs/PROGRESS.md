# Progress

Last updated 2026-11-03.

## Done (verified)

| What | Evidence |
|---|---|
| 15 spreadsheet cases with inputs, tasks (en/zh), rules, reference, failure artifact, limitations | `cases/`, `failbook verify --seeds 3` (CI) |
| Checker validation: reference + alternatives pass (also after a LibreOffice re-save), failure fails on the named rules, review-found wrong answers (`fail_*.py`) fail, untouched input fails | `failbook verify` output in CI |
| Core claim: live, preserve and structure rules catch the failures that value checks miss, without rejecting correct solutions <!-- sync-numbers --> | [evidence/checker-ablation.md](evidence/checker-ablation.md) |
| Evidence chain for one case, end to end, no AI needed | [evidence/demo-xlsx-001.md](evidence/demo-xlsx-001.md), `scripts/demo-evidence-chain.sh` |
| Fix with a deterministic self-check: flags 13/15 failures, no error on 41 correct solutions | [evidence/fix-self-check.md](evidence/fix-self-check.md) |
| Harbor export loads in Harbor 0.24.0; SkillEvaluator 0.7.0 `tier3 validate --harbor-contract` passes | [DECISIONS.md](DECISIONS.md) D4 |
| Static site: catalog + evidence pages, en/zh, reflow at 320 px checked | `failbook site`, GitHub Pages workflow |
| Audit regression corpus: 361 labelled probes (317 before/after pairs, 44 single files) from the reviews, run with the tests | [tests/audit_corpus/](../tests/audit_corpus/README.md) |
| Two independent adversarial reviews (52 then 33 findings; 51 and 28 confirmed and fixed), each fix re-checked until the re-checks reported only stated limitations | [DECISIONS.md](DECISIONS.md) D9, D10; `alt_*.py` / `fail_*.py` fixtures |

## Release steps left to the maintainer

- Merge the first release (pull request #1 on Ray0766/failbook).
- Enable GitHub Pages with "GitHub Actions" as the source (Settings → Pages); the `pages` workflow then
  publishes the case gallery on every push to `main`. The build environment could not change this setting.

## Not done / not verified

- **No recorded agent runs.** Every case is `reconstructed`. Running a CLI agent unsandboxed was blocked in
  the build environment, and no Docker daemon was available for Harbor. Next: run
  `failbook run <case> --agent <preset> --model <name>` (as a non-root user) or a Harbor job, and commit the
  result files.
- **The fix's effect on agents is a hypothesis.** Next: the same runs with `--fix spreadsheet-verify-skill`,
  or `failbook export skillevaluator` and SkillEvaluator Tier 3 for a with/without comparison.
- Inspect path (via `inspect-harbor`) not tested.
- The audit script was tuned on these 15 failures and on the reviewers' probe workbooks; it needs new failures
  to test generalisation. Its known gaps are listed in the skill's SKILL.md ("What the audit script cannot tell").
- The audit corpus in the repository holds 361 hand-made, labelled probes; the 12 large timing pairs and the
  random families used during the reviews are not in it, and it does not compare with earlier versions.
- The conditional-format, chart and number-format checks were compared with LibreOffice only, not with Excel.
- No full accessibility audit of the site.
