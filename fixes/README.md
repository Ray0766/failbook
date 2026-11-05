# Fixes

A fix is a remedy you can retest: an [Agent Skill](https://agentskills.io/) (`kind: skill`, a `skill/`
folder with `SKILL.md`) or a prompt addition (`kind: prompt`, a `prompt.md`). Each has a `fix.yaml` with
`id`, `title`, `status` (`candidate` until results show an effect) and a description.

| Fix | Status | What it is | Evidence so far |
|---|---|---|---|
| [spreadsheet-verify-skill](spreadsheet-verify-skill/) | candidate | Checklist of the pitfalls behind the cases plus `scripts/audit_xlsx.py`, a task-agnostic audit of an edited workbook against its original | The audit flags 13/15 failure artifacts and raises no error on 41 correct solutions ([docs/evidence/fix-self-check.md](../docs/evidence/fix-self-check.md)). **No agent runs yet**, so its effect on agents is unknown. |

Retest a fix:

```bash
failbook run xlsx-011 --agent claude-code --model <model> --attempts 3                                 # baseline
failbook run xlsx-011 --agent claude-code --model <model> --attempts 3 --fix spreadsheet-verify-skill  # with the fix
# or with/without lift through NVIDIA SkillEvaluator (Docker):
failbook export skillevaluator --fix spreadsheet-verify-skill --out dist/se
skillevaluator tier3 evaluate dist/se/spreadsheet-verify --agents claude-code
```

For Claude Code the skill is installed as a project skill (`.claude/skills/`). Other agents get the whole
skill folder, scripts included, in `skills/<skill-name>/` of their working directory, plus an instructions
file (`AGENTS.md`, `GEMINI.md`) that contains the skill text and points at that folder. Run the baseline and the fix with the
same agent version, model, seed and number of attempts.
