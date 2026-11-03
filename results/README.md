# Results

One JSON file per attempt, in `results/<case-id>/`. Files are **append-only**: a retest is a new file, and
nobody edits or deletes an existing one (a reviewer's classification goes in a follow-up commit that only
fills `classification.reviewed` and `classification.note`).

## Producing a result

```bash
# Failbook drives the agent (not sandboxed: run as a normal user, in a throwaway account or VM)
failbook run xlsx-004 --agent claude-code --model <model> --attempts 3 --submitter <github-handle>

# Any other tool (a chat app, an IDE agent, a Harbor job): prepare, let the agent work, then record
failbook prepare xlsx-004 /tmp/try && <agent works on /tmp/try using TASK.md>
failbook grade xlsx-004 /tmp/try --record results --agent cursor --agent-version 2.3 --model <model> \
    --harness manual --final-message reply.txt --submitter <github-handle>
```

Submit **every** attempt, including timeouts and crashes; do not pick the best one. Then run
`failbook docs` and open a pull request.

## Schema `failbook.result/1`

| Field | Meaning |
|---|---|
| `case`, `case_version`, `seed`, `lang` | which variant of which case version |
| `agent` `{name, version, command}`, `model`, `fix`, `harness` | what ran, and how (`failbook-run`, `harbor`, `manual`) |
| `attempt`, `attempts`, `started_at`, `duration_sec` | |
| `run_status` | `completed`, `timeout`, `agent-error`, `not-run` |
| `verdict`, `score`, `rules[]` | grader output; `rules[].status` is `pass` / `fail` / `inconclusive` |
| `classification` `{auto, reviewed, note}` | see [docs/FAILURE_TAXONOMY.md](../docs/FAILURE_TAXONOMY.md) |
| `engine`, `failbook_version` | grader provenance |
| `prompt_sha256`, `inputs`, `outputs` | SHA-256 of the prompt, the input files and the files left behind |
| `environment` `{os, python, isolation}`, `budget` `{timeout_sec, cost_usd, tokens}` | `null` means unknown, not zero |
| `final_message`, `claimed_done` | the agent's last reply (truncated) and, after review, whether it claimed success |
| `submitter`, `notes` | |

`failbook report` (also run by `failbook docs`) turns these files into [RESULTS.md](../RESULTS.md).
