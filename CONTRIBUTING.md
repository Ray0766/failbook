# Contributing

Thank you! Read [AGENTS.md](AGENTS.md) and [docs/QUALITY_BRIEF.md](docs/QUALITY_BRIEF.md) first: they set
the bar for evidence (facts vs hypotheses, no invented results, counts not rates).

## Set up

```bash
git clone https://github.com/Ray0766/failbook && cd failbook
python3 -m venv .venv && . .venv/bin/activate
pip install -e ".[dev]"
# LibreOffice is required for grading: apt install libreoffice-calc / brew install --cask libreoffice
failbook verify xlsx-001          # checks your setup end to end (about 10 s)
pytest -q
```

## Ways to contribute

- **Submit retest results.** See [results/README.md](results/README.md). Submit every attempt.
- **Propose a case.** Open a "New case" issue with the task, a sample input and what goes wrong. Synthetic
  or publicly shareable files only. To implement it, follow [docs/WRITING_CASES.md](docs/WRITING_CASES.md);
  `failbook verify <id> --seeds 3` must pass.
- **Dispute a checker.** If a correct output is graded `fail` (or a wrong one `pass`), open a "Checker
  dispute" issue with the output file. Confirmed problems are fixed, an alternative solution is added as a
  regression test, and the case version is bumped.
- **Propose a fix.** Add `fixes/<id>/` (see [fixes/README.md](fixes/README.md)) and, ideally, with/without results.

## Before you open a pull request

```bash
failbook verify <changed cases> --seeds 3 --check-committed
failbook docs                      # regenerates case READMEs, the index and RESULTS.md
pytest -q
```

Case code (`case.py`, solutions) is executed by the grader. Pull requests that add or change it are
reviewed for anything beyond reading and writing workbooks; CI runs them without secrets.
