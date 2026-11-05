# Audit regression corpus

Before/after workbook pairs that pin how the fix's audit script (`src/failbook/audit.py`, shipped as
`fixes/spreadsheet-verify-skill/skill/scripts/audit_xlsx.py`) treats edits that the reviews of the
script found hard: correct edits it must not flag, wrong edits it must flag, and files it must refuse.
[`tests/test_audit_corpus.py`](../test_audit_corpus.py) audits every pair and fails when an outcome
leaves its label.

Why it exists: the audit's rules are heuristics, and during the reviews
([docs/DECISIONS.md](../../docs/DECISIONS.md) D7, D9) several fixes traded one false finding for another.
From the third fix round on, every change was accepted only if none of these pairs got worse. Keeping
them here lets anyone who changes the audit run the same check.

## What is in it

- [`probes.yaml`](probes.yaml): 273 probes (229 before/after pairs and 44 single files). Each has an `id`, a `label`, what the `edit` did, `why` that
  label is the intended outcome, the `after` workbook and, for most, the `before` (original) workbook.
- `workbooks/`: the 433 files they use (about 2.4 MB), grouped by the round that built them. Several
  probes share an original.

All files are synthetic: workbooks written with openpyxl from made-up figures (three re-saved by
LibreOffice 24.2), and a few deliberately broken files (empty, truncated, not a zip, no workbook part)
that the audit must refuse. They hold no personal data and are licensed like the rest of the
repository (MIT).

## Labels

The outcome of an audit without recalculation (`recalc=False`; the `recalculation-skipped` notice is
ignored):

| Label | Passes when |
|---|---|
| `error` | at least one error-level finding |
| `warning` | warnings only |
| `clean` | no finding |
| `error\|warning` | at least a warning (a wrong edit the audit may only warn about) |
| `warning\|clean` | no error (a correct edit; a warning is tolerated) |
| `unreadable` | the audit refuses the file (exit code 3 on the command line) |
| `any` | it does not crash (a known trade-off, kept for information) |

Other fields:

- `require` / `forbid`: regular expressions that the findings' text must or must not match, for
  messages a review found false.
- `known_limit: true`: a limitation listed in the skill's SKILL.md ("What the audit script cannot
  tell"). The label says what would be right. When today's outcome is outside it, `now` records that
  outcome and the test pins it: if it changes, either the limitation was fixed (remove `now` and
  `known_limit`, and update SKILL.md) or it got worse. `now: require-unmet` marks a limitation where
  the outcome is right but the findings miss the cell the label is about: the test checks that the
  `require` pattern still does not match. A probe whose outcome is pinned by any other `now` has no
  `require` or `forbid`, because they are not checked.
- `libreoffice_sample: true`: also audited with LibreOffice recalculation, when LibreOffice is
  installed; `with_libreoffice` gives the label for that run when it differs.
- `source`: the round that built the probe, as numbered in DECISIONS.md D9: `review-2` (the second
  independent review), then `fix-1`, `recheck-1`, `fix-2`, `recheck-2`, `fix-3`, `recheck-3` and `fix-4`
  (each fix round and the skeptical re-check that followed it), `check` for the probes that the
  independent check of this corpus added, and `misses` for those that pin the fix of the misses the
  corpus first recorded as limitations, and the false findings its checks found; `position` for
  references meant to read a position (the top row) after a sort. The id prefix names the same round.

## Running it

```bash
pytest -q tests/test_audit_corpus.py          # about 4 s; the LibreOffice sample adds about 20 s
pytest -q tests/test_audit_corpus.py -k recheck2-i4
```

A failure prints the probe's edit, the outcome, the label and the findings.

## When a probe fails

- If your change to the audit caused it, the change brings back a false finding or loses a true one
  that a review established: fix the change.
- Change a label only with evidence that it is wrong (what Excel and LibreOffice show, or what the task
  allows), and say why in the commit message.

## Adding a probe

Build the pair with a short openpyxl script (or save it from a spreadsheet application), put the files
under `workbooks/<set>/`, and add an entry with `id`, `label`, `edit`, `why`, `after`, `before` and
`source` (a new round name also goes into `SOURCES` in the test, with its id prefix).
`test_probes_are_well_formed` checks that every stored file is used, that every entry has the required
fields and no unknown or repeated key, and that every pattern compiles.

## What is not here

- 12 large pairs (12.9 MB) that the reviews used to time the audit on 5,000 to 20,000 rows.
- Three families of randomly generated pairs (1,370 in all) used during the reviews.
- A comparison with earlier versions of the audit: the reviews ran one; this test checks the labels.
- Excel: outcomes were judged against what LibreOffice shows and against the tasks, not checked in Excel.
