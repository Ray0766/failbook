# Datasheet: Failbook spreadsheet cases v0.1

Structured after *Datasheets for Datasets* (Gebru et al., arXiv:1803.09010, revised 2021), a classic
method reference, not a recent result.

## Motivation

- **Purpose.** Reproducible cases where an AI agent reports that a spreadsheet task is done while the file is
  wrong, with deterministic checks that anyone can rerun on an agent's output.
- **Created by.** The Failbook maintainers, November 2026. No funding.

## Composition

- 15 cases (`cases/xlsx-001` to `xlsx-015`), each: a task in English and Chinese, input workbook(s), 3 to 7
  rules, a reference solution, a failure reproduction (`failure.py`) with its artifact and grading, optional
  alternative solutions, limitations.
- **Data.** Synthetic and fictional: names, companies and products come from short lists in
  `src/failbook/kit.py`; numbers are random. No personal data. Every input is regenerated from a seed;
  seed 0 is committed.
- **Failure artifacts.** All 15 are reconstructions (evidence level `reconstructed`): the output of running a
  naive procedure, not a recording of a model. No case is `observed` yet.
- **Not included.** Real agent transcripts, model outputs, or any measurement of how often agents fail.

## Collection process

- Failure patterns were chosen because they are commonly reported in spreadsheet automation and are
  silent (the file opens and looks plausible). The selection is the authors' judgement; there is no sampling
  frame. **Selection bias is by design**: the set over-represents subtle failures.

## Preprocessing and validation

- Inputs are written with openpyxl and then recalculated and re-saved by LibreOffice, so they carry cached
  values like a file saved by a spreadsheet application.
- Checker validation per case and seed (CI runs seeds 0 to 2): reference and alternatives pass, failure fails
  on the named rules, the untouched input fails, the reference passes after a LibreOffice re-save.

## Uses

- **Intended.** Reproducing and explaining these failure patterns; retesting agents, prompts and skills on
  them; testing graders (see `docs/evidence/checker-ablation.md`).
- **Not intended.** Ranking agents or estimating general reliability: pass counts over a curated set say
  nothing about performance on real work. Training data: the tasks and solutions are public, so treat any
  model trained after November 2026 as possibly contaminated (use non-zero seeds to reduce the effect).

## Distribution and licence

- MIT licence (code, cases, data). Synthetic data only.

## Maintenance

- Cases are versioned (`version` in `case.yaml`); a change to a task or a rule bumps the version, and results
  for older versions are kept but shown separately.
- Results are append-only. Disputes about a checker are handled through the "checker dispute" issue form;
  a confirmed checker error is fixed and the case version bumped.
- A case that every recorded agent configuration passes consistently may be marked `retired`; it stays in
  the repository.
