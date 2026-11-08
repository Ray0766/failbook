# Failbook: primary-source verification of technical claims

Checked 2026-11-03. Every finding below comes from a primary artifact that was fetched or read during this session: official source repositories through raw.githubusercontent.com, the PyPI JSON API, the npm registry, the Ubuntu archive, code.claude.com, or the installed package source. Some findings were also run locally on this machine (Ubuntu 24.04, LibreOffice 24.2.7.2 from apt).

**Hosts blocked by the session network policy** (each returned a DNS failure or a proxy 403): openpyxl.readthedocs.io, foss.heptapod.net, wiki.documentfoundation.org, www.libreoffice.org, help.libreoffice.org, bugs.documentfoundation.org, packages.debian.org, deb.debian.org (and every Debian mirror tried), security-tracker.debian.org, harborframework.com, docs.harborframework.com, inspect.aisi.org.uk, meridianlabs-ai.github.io, developers.openai.com, geminicli.com, formulas.readthedocs.io, and github.com (web and API). Where one of these was blocked, I used the same content from another primary channel, as noted in each section.

Confidence scale: **High** means verified in source or docs, or run locally. **Medium** means the primary page was blocked and the finding rests on search-index snippets or press coverage.

---

## 1. openpyxl

### 1a. insert/delete/move do not update formulas or references. Verified: High

The official docs (3.1.5 doc sources) say, in `editing_worksheets.rst`, lines 28-34:

> "Openpyxl does not manage dependencies, such as formulae, tables, charts, etc., when rows or columns are inserted or deleted. This is considered to be out of scope for a library that focuses on managing the file format. As a result, client code **must** implement the functionality required in any particular use case."

Lines 47-51 of the same file cover `move_range`:

> "If cells contain formulae you can let openpyxl translate these for you, but as this is not always what you want it is disabled by default. Also only the formulae in the cells themselves will be translated. References to the cells from other cells or defined names will not be updated"

- **Source.** The canonical URL is https://openpyxl.readthedocs.io/en/stable/editing_worksheets.html, which was blocked. I read the same doc sources from Ubuntu's build of the upstream docs: http://archive.ubuntu.com/ubuntu/pool/universe/o/openpyxl/python-openpyxl-doc_3.1.5+dfsg-4_all.deb, at `/usr/share/doc/python-openpyxl-doc/html/_sources/editing_worksheets.rst.txt`. The changelog in that package starts at "3.1.5 (2024-06-28)".
- **Source code, 3.1.5** (`openpyxl/worksheet/worksheet.py`):
  - `insert_rows`, `insert_cols`, `delete_rows` and `delete_cols` call `_move_cells`, which calls `_move_cell(row, column, row_offset, col_offset)` with `translate=False`.
  - The `move_range` docstring says "Formulae and references will not be updated." Its `translate=True` option only runs `Translator` on the moved cells themselves.
- **Local test.** Sheet: A1=1, A2=2, A3=`=SUM(A1:A2)`, B1=`=A2*10`. After `insert_rows(2)`, B1 is still `=A2*10` (which now points at an empty cell) and A4 is `=SUM(A1:A2)`, so the range no longer covers the shifted value.
  - `move_range("A1:B1", rows=2, translate=True)` changed B3 to `=A3*2`, but C1 stayed `=B1+1` and now points at the vacated cell.
- **Implication for Failbook.** This failure mode is real and documented, so a "row inserted, formulas silently broken" case is well founded. Graders should check that dependent formulas point at the right cells, not only that the values are present.

### 1b. Charts and images on load/save. Verified: High

The current official doc warns only about **shapes**. `tutorial.rst`, "Loading from a file", lines 377-381:

> "openpyxl does currently not read all possible items in an Excel file so shapes will be lost from existing files if they are opened and saved with the same name."

- **Older docs.** The older text, "images and charts will be lost", appears in old doc versions (2.4 and similar, per search results). It is not in the 3.1.5 docs.
- **Source code, 3.1.5:**
  - `reader/excel.py` lines 265-269 (worksheets) and 205-209 (chartsheets) call `find_images()` and then `ws.add_chart(c, c.anchor)`.
  - `reader/drawings.py::find_images` reads charts through `read_chart(ChartSpace)`.
  - Images are read only when Pillow is installed: `if not PILImage: # Pillow not installed, drop images`. WMF images are always dropped.
  - Unsupported DrawingML produces the warning "DrawingML support is incomplete and limited to charts and images only. Shapes and drawings will be lost."
- **Local test.** A bar chart survived a `load_workbook` → `save` round trip (`chart1.xml` and `drawing1.xml` were present after the second save). Pillow was not installed in the test venv, so existing images would have been dropped.
- **Inference, not verified.** Charts are re-serialized through openpyxl's own object model, so chart XML that the model does not cover may still be lost.
- **Implication for Failbook.** Do not claim "openpyxl drops charts". The accurate claims are:
  - Shapes and drawings are dropped.
  - Images are dropped when Pillow is missing, and WMF images are always dropped.
  - Charts are rebuilt and may lose fidelity.

### 1c. Latest release. Verified: High

openpyxl **3.1.5**, uploaded 2024-06-28 (https://pypi.org/pypi/openpyxl/json). Nothing has been uploaded since, pre-releases included. Ubuntu noble ships 3.1.2+dfsg-6.

---

## 2. LibreOffice

### 2a. OOXMLRecalcMode. Verified: High

The official schema is `officecfg/registry/schema/org/openoffice/Office/Calc.xcs`, at https://raw.githubusercontent.com/LibreOffice/core/master/officecfg/registry/schema/org/openoffice/Office/Calc.xcs. The key sits at `group Formula` → `group Load` → `prop OOXMLRecalcMode` (`xs:int`), lines 1607-1634 on master:

- desc: "Specifies whether to force a hard recalc after load on OOXML-based Excel documents (2007 and newer)."
- `0` = "Recalc always", `1` = "Recalc never", `2` = "Ask before Recalc".
- **The default is `<value>1</value>` (Recalc never).**

The `libreoffice-24-2` branch has the same meanings and the same default. The installed 24.2.7 `share/registry/main.xcd` also has default `1`.

A sibling key, `ODFRecalcMode`, uses the same 0/1/2 scheme for ODF files.

**Local test** (LibreOffice 24.2.7.2): I took an xlsx with stale cached results (B1 changed from 1 to 100 in the XML, D4 `=SUM(B1:B3)` still cached as 6) and converted it headless with `/org.openoffice.Office.Calc/Formula/Load/OOXMLRecalcMode`:

| Mode | D4 after conversion |
| --- | --- |
| 0 | 105 (recalculated) |
| 1 | 6 (stale cache kept) |

**Implication for Failbook.** The recalculating grader must set the mode to 0 explicitly. The default (1) trusts cached values, which is exactly the result an agent can get wrong.

### 2b. XLOOKUP introduced in 24.8. Verified: High

Release notes wiki: blocked. Two other primary sources confirm it:

- **Official help source.** https://raw.githubusercontent.com/LibreOffice/help/libreoffice-24-8/source/text/scalc/01/func_xlookup.xhp ("XLOOKUP Function") embeds `text/scalc/00/avail_release.xhp#24.8`. That section reads: "This function is available since %PRODUCTNAME 24.8." The same file returns 404 on the `libreoffice-24-2` branch.
- **Core source.** In `include/formula/opcode.hxx`, `ocXLookup`, `ocXMatch`, `ocFilter` and `ocLet` are present on `libreoffice-24-8` and absent on `libreoffice-24-2`.

Press and search coverage (Phoronix, 2024-08-22) says 24.8 was released on 2024-08-22 with XLOOKUP and XMATCH. Confidence on that date: Medium.

**Local test, 24.2.7.2.** `=_xlfn.XLOOKUP(...)` gives `#NAME?`, and the round trip rewrites it in lowercase as `_xlfn.xlookup`. TEXTJOIN and SUBTOTAL work in 24.2.

### 2c. Current releases and distro versions

**LibreOffice upstream.** Confidence: Medium for release dates, High for tag existence.

- Tags present in the core repo: `libreoffice-26.8.2.2`, `libreoffice-26.2.7.2`, `libreoffice-25.8.7.2`.
- Branch versions in `configure.ac`:

  | Branch | Version |
  | --- | --- |
  | `libreoffice-26-8` | 26.8.3.0.0+ |
  | `libreoffice-26-2` | 26.2.8.0.0+ |
  | `master` | 27.2.0.0.alpha0+ |

- Press (libreoffice.org blocked):
  - 26.8 was released 2026-08-26.
  - **26.8.2** was announced 2026-10-29.
  - 26.2.6 shipped 2026-09-04, and 26.2.7 was scheduled for 2026-10-29.
  - The 26.2 series reaches end of life on 2026-11-30.
- **As of early November 2026:** 26.8.2 is the newest release, and 26.2.7 is the older series, supported until 2026-11-30.

**Ubuntu 24.04 (noble).** Verified: High.

- Release: `libreoffice-calc 4:24.2.2-0ubuntu1`.
- noble-updates and noble-security: **`4:24.2.7-0ubuntu0.24.04.7`**.
- Source: http://archive.ubuntu.com/ubuntu/dists/noble-updates/main/binary-amd64/Packages.xz, Release file dated 2026-11-02.
- This machine also has 24.2.7.2 installed. **Ubuntu 24.04 therefore cannot evaluate XLOOKUP.**

**Debian 13 (trixie).** Medium: packages.debian.org and every Debian mirror were blocked, so this rests on search-index snippets from packages.debian.org, piuparts.debian.org and the security tracker.

- Trixie ships **LibreOffice 25.2.3** (`4:25.2.3-2`), with point and security revisions `+deb13uN` (snippets show deb13u4 to deb13u8).
- sid has moved to 26.2.x.
- Since 25.2 is newer than 24.8, XLOOKUP should work on trixie. That is an inference.

### 2d. libreoffice-calc-nogui in trixie. Medium (Debian), High (Ubuntu)

- **Debian.** Search-index snippets show `packages.debian.org/sid/libreoffice-calc-nogui` ("office productivity suite -- spreadsheet (no GUI variant)") and piuparts logs for `libreoffice-calc-nogui_4:26.2.2.2-3+b2` (stable2sid). The trixie `libreoffice-nogui 4:25.2.3-2+deb13u4` metapackage depends on it. I could not open the trixie page directly.
- **Ubuntu noble** (same Debian-derived packaging). `libreoffice-calc-nogui 4:24.2.7-0ubuntu0.24.04.7` exists in universe and `Conflicts: libreoffice-calc`. It depends on `libreoffice-core-nogui | libreoffice-core`.
- **Implication for Failbook.** A Debian trixie image with `libreoffice-calc-nogui` gives a headless LibreOffice 25.2 that handles XLOOKUP. An `ubuntu:24.04` base image does not.

---

## 3. Harbor (harbor-framework/harbor)

### Versions and docs. High

- PyPI (https://pypi.org/pypi/harbor/json):
  - Latest stable: **0.25.0, 2026-10-25**.
  - The installed copy is **0.24.0, 2026-10-05**.
  - Nightly builds such as `0.25.1.devYYYYMMDD` continue through 2026-11-02.
  - Author: Alex Shaw.
- **Docs.** The PyPI README badge points to https://harborframework.com/docs. The repo's `docs/README.md` says the documentation now lives in `docs-mintlify/` and is published at **https://docs.harborframework.com**, with redirects kept from the old URLs. Both hosts were blocked, so I read the Mintlify sources from raw GitHub.
- **Important for SkillEvaluator.** SkillEvaluator's `pyproject.toml` pins `harbor[gke]==0.24.0`.

### Task layout. High

From `harbor/models/task/paths.py` (0.24.0) and `docs-mintlify/tasks/overview.mdx`:

```
instruction.md, task.toml, environment/ (Dockerfile | docker-compose.yaml | ...),
solution/solve.sh, tests/test.sh   (+ optional trajectory.json, steps/<name>/...)
```

- In the container: tests go to `/tests` and the solution to `/solution`.
- Logs go to `/logs/{agent,verifier,artifacts}`.
- Task skills are placed through `environment.skills_dir` (docs: `tasks/skills.mdx`).

### task.toml schema_version. High

- `TaskConfig.schema_version: str = "1.4"` in `models/task/config.py` line 802. The value is identical on GitHub `main` and tag `v0.25.0`.
- The field is a plain string with no validation of its value.
- A legacy top-level `version` key is renamed to `schema_version` (`handle_version_rename`).
- The docs example (`tasks/configuration.mdx`) starts with `schema_version = "1.4"`.

### Reward contract. High

From `harbor/verifier/verifier.py` (0.24.0, same check on main and v0.25.0):

- Paths are `/logs/verifier/reward.json` and `/logs/verifier/reward.txt`. **`reward.json` wins if both exist.**
- `reward.txt` must contain one finite float, which becomes `{"reward": value}`.
- In `reward.json`, **every top-level value must be an `int` or `float` and must be finite.** Any other value raises `VerifierOutputParseError` ("Non-numeric reward ..."). That includes strings, nested objects, lists and NaN.
- An empty file raises `RewardFileEmptyError`. If neither file exists, `RewardFileNotFoundError` is raised.
- Docs (`tasks/verifier.mdx`): "`/logs/verifier/reward.json` | JSON key-value object of labeled numeric metrics" and "If both are present, Harbor prefers `reward.json`."

### CLI usage. High

`harbor run --help` (0.24.0) confirms these flags:

| Flag | Meaning |
| --- | --- |
| `-p/--path` | Path to a local task or dataset directory |
| `-a/--agent` | Agent name, or an import path `module:Class` |
| `-m/--model` | Model; repeatable |
| `-k/--n-attempts` | Number of attempts per trial |
| `-n/--n-concurrent` | Concurrent trials; default 4 |
| `-e/--env` | Environment; default docker |

The docs (`jobs/run-a-job.mdx`) show `harbor run -p "<path>" -a "<agent>" -m "<model>" -n ... -k ... -r ...`. **So `harbor run -p <path> -a <agent> -m <model> -k <n>` is valid.**

### Built-in agents. High

From the `AgentName` enum in `harbor/models/agent/name.py`, 0.24.0:

oracle, nop, acp, claude-code, cline-cli, cortex-code, terminus, terminus-1, terminus-2, aider, codex, cursor-cli, gemini-cli, antigravity-cli, rovodev-cli, goose, grok-build, hermes, junie, mini-swe-agent, nemo-agent, swe-agent, opencode, openclaw, openhands, openhands-sdk, antigravity-sdk, kimi-code, kimi-cli, langgraph, deerflow, mimo, mcode, pi, qwen-coder, copilot-cli, code-puppy, devin, trae-agent, computer-1, eve, fx, fx-dev, dspy-rlm, vibe, muse-code, strands.

How Harbor invokes the three CLIs (0.24.0 source):

- **claude-code:** `claude --verbose --output-format=stream-json --permission-mode=bypassPermissions ... --print`, with the prompt on stdin and `IS_SANDBOX=1`.
- **codex:** `codex exec --dangerously-bypass-approvals-and-sandbox --skip-git-repo-check --model <m> --json ...`.
- **gemini-cli:** `gemini --yolo --model=<m> --prompt=<instr>`, with `GEMINI_CLI_TRUST_WORKSPACE=true`.

**Implication for Failbook.** The Failbook export should write a flat numeric `reward.json`, for example `{"overall": 0.0, "check_x": 1.0}`, and keep diagnostics in a separate file under `/logs/verifier/`.

---

## 4. NVIDIA SkillEvaluator. High

**Version.**

- `pyproject.toml`: **0.7.0**. Local clone at commit e4cda13, dated 2026-11-02 23:24 UTC, tag `ci-0.7.0`.
- PyPI `skillevaluator` 0.7.0 was uploaded 2026-11-02T20:08Z.
- `CHANGELOG.md` has no 0.7.0 heading yet: the latest dated heading is "## 0.6.0 - 2026-10-20", followed by "Unreleased".

**Tier 3 and native Harbor tasks.**

- `docs/tier3-live-evaluation.mdx` describes "Run a real agent with and without your skill in a sandboxed environment, then measure Skill Lift". It also says "**Skill Lift** is the signed difference between the two arms".
- Outputs include `lift.json`, `with-skill/` and `without-skill/`.
- `--skip-baseline` drops the lift.
- `docs/custom-graders.mdx` ("Bring your own task"): "a native Harbor task you own under `evals/harbor/`". The scaffold command is `skillevaluator init-harbor-task`, whose `--mode` defaults to `custom_only`.
- `docs/eval-datasets.mdx`: `harbor.task_source` can be `auto`, `evals_json` or `native_harbor`.

**custom_only and `overall`.** From `custom-graders.mdx`:

> "In `custom_only` mode, a numeric `overall` in `reward.json` or a numeric `reward.txt` is required."

> "In `custom_only` mode, your `overall` reward drives pass@k and the overall lift."

> "Harbor 0.24 requires every top-level value in its `reward.json` to be a number, for example `{"overall": 1.0, "native_quality": 1.0}`; SkillEvaluator reports each numeric key other than `overall` as a custom metric. Nested objects such as `custom_metrics` or `details` make Harbor reject the reward with `VerifierOutputParseError`, so use `tests/grader.py` when you need them."

`custom_lift` and `custom_lift.json` hold the per-metric with-skill minus without-skill difference.

**Supported agents** (`agents-and-sandboxes.mdx`): only `claude-code` (alias `claude`), `codex` and `opencode`. **Gemini CLI is not supported.** Harbor installs pinned versions: Claude Code 2.1.292 and Codex 0.161.0.

**Implication for Failbook.**

- Export tasks with `custom_only` and a flat `reward.json` containing `overall`.
- Stay on Harbor 0.24.0 to match SkillEvaluator's pin.
- Present a Gemini comparison as Harbor-only, not as SkillEvaluator.

---

## 5. Inspect (UK AISI) and inspect_swe. High (docs read from repo sources; the website was blocked)

### Versions (PyPI)

| Package | Latest | Date | Notes |
| --- | --- | --- | --- |
| `inspect-ai` | **0.3.287** | 2026-10-31 | Docs: https://inspect.aisi.org.uk/ |
| `inspect-swe` | **0.2.72** | 2026-10-20 | meridianlabs-ai/inspect_swe |
| `inspect-harbor` | **1.0.0** | 2026-09-23 | "Inspect AI interface to Harbor tasks", meridianlabs-ai/inspect_harbor |

### Sandboxing

Source: `docs/sandboxing.qmd` on main.

- **Built-in providers:** `docker` and `local` (no sandbox).
- **External packages:** k8s, openshell (NVIDIA OpenShell) and others.
- **Per-sample sandboxes:** "each sample gets its own sandbox *instance*".
- **Network:** "The generated Compose configuration sets `network_mode: none`". A custom compose file *replaces* that setting, so you must add `network_mode: none` yourself.
- **Resource limits** go in `compose.yaml` or `ComposeConfig(... mem_limit="512m", cpus=..., network_mode="none")`.
- **Output limits:** 100MB file reads and 10MB exec output by default. `exec()` output is silently front-truncated.
- **Scorers** can call `sandbox().exec()` and `read_file()`. Scorer code itself runs on the host, so a deterministic Python and LibreOffice grader can run inside the container through `sandbox().exec()`.

### CLI agents

- **inspect_swe** (docs `index.qmd`) "makes software engineering agents like Claude Code, Codex CLI, Gemini CLI, Kimi Code, OpenCode, and Mini SWE Agent available as standard Inspect agents".
- **Exports:** `claude_code`, `codex_cli`, `gemini_cli`, `kimi_code`, `opencode`, `mini_swe_agent`, `antigravity`, plus `interactive_*` ACP variants.
- **Options** include `skills`, `attempts`, `cwd`, and `version` (`auto`, `sandbox`, `stable`, `latest`, or a pinned `x.x.x`).
- **Model bridge.** The agents run inside the sandbox, and their model calls go through `sandbox_agent_bridge()` to a proxy on the sandbox loopback (localhost:13131). So they work under `network_mode: none` and with any Inspect model.
- **Skills.** Inspect also has a generic `skill()` tool that copies skills into the sandbox (`tools-standard.qmd`).

### Running Harbor tasks inside Inspect

- inspect_harbor loads local Harbor tasks or datasets with `harbor(path=...)`.
- Its `harbor_scorer` reads `/logs/verifier/reward.json` or `reward.txt`.
- It can be combined with `--solver inspect_swe/claude_code`.
- Caveat from its docs: "Inspect fails the whole eval on the first sample error by default, whereas Harbor continues the job."

### Log Viewer

Source: `log-viewer.qmd`.

- Opens with `inspect view` or the VS Code extension.
- Has a live view during runs and per-sample Messages, Scoring and Metadata tabs.
- Supports filtering and sorting, and reads S3 logs.
- `inspect view bundle` publishes a static site (GitHub Pages, S3, HF Spaces).
- Harbor has `harbor view` ("Start web server to browse trajectory files").

### Fair comparison

| Aspect | Harbor | Inspect |
| --- | --- | --- |
| Task format | Task directory and shell verifier, framework-independent; SkillEvaluator consumes it | Python `Task`/`Scorer` objects; Harbor tasks also run through inspect_harbor |
| Agent network | Agents call providers directly; `public`/`allowlist`/`no-network` per phase | Model proxy lets agents run with no container network |
| Multiple runs | `-k` attempts | `--epochs` plus reducers (`pass_at_k`) |
| Built-in CLI agents | ~45 | 7 through inspect_swe |
| Log viewer | `harbor view` | Richer, bundleable |
| Scorer | Any shell script writing `reward.json` | Python scorer calling `sandbox().exec` |

The custom deterministic grader is portable either way.

---

## 6. Python formula engines. High

| | Latest (PyPI) | XLOOKUP | TEXTJOIN | SUBTOTAL |
|---|---|---|---|---|
| **pycel** | **1.0b30**, 2021-10-13 (pypi.org/pypi/pycel/json) | **No**: `# def xlookup(value):` stub commented out in `lib/lookup.py`, also on master | **No**: commented-out stub in `lib/text.py` | **Yes**: `func_subtotal` in `excelformula.py`, codes 1-11 and 101-111. The source notes "This does not implement skipping hidden rows." |
| **formulas** | **1.3.4**, 2026-03-11 (pypi.org/pypi/formulas/json) | **Yes**: `FUNCTIONS['_XLFN.XLOOKUP']`; CHANGELOG v1.3.1 (2025-09-15) "Add `XMATCH` and `XLOOKUP`" | **Yes**: CHANGELOG v1.2.9 (2025-04-05) "Add `FILTER`, `TRANSPOSE`, `SUBSTITUTE`, `TEXTJOIN`, `T`" | **No**: not in 1.3.4 source or on master |

Local evaluation of the same workbook:

- pycel: `UnknownFunction` "Function XLOOKUP is not implemented" (and the same for TEXTJOIN); SUBTOTAL(9) = 6.
- formulas: XLOOKUP = 2, TEXTJOIN = "a,b,c", SUBTOTAL = `#NAME?`.

pycel's README gives no function list. It says "I have only added support for functions that I need".

**Implication for Failbook.** Neither engine covers all three functions. Keep LibreOffice 24.8 or later as the reference engine, and use pycel or formulas only as optional cross-checks.

---

## 7. Agent CLI non-interactive flags

### 7a. Claude Code. Verified: High

Latest npm release: 2.1.317, 2026-11-02. The sources are https://code.claude.com/docs/en/cli-reference.md, `skills.md`, `permission-modes.md` and `headless.md`.

- `--print`, `-p`: "Print response without interactive mode".
- `--dangerously-skip-permissions`: "Skip permission prompts. Equivalent to `--permission-mode bypassPermissions`."
- `--model`: "Sets the model for the current session with a model alias ... or a model's full name".
- The docs show `claude -p "<prompt>" --dangerously-skip-permissions` explicitly, for "Run fully unattended inside a container".
- **Caveat** (`permission-modes.md`): "On Linux and macOS, Claude Code refuses to start in `bypassPermissions` mode when running as root or under `sudo`". The check is skipped inside a recognized sandbox; Harbor sets `IS_SANDBOX=1`.
- **Skills** (`skills.md`): "Project | `.claude/skills/<skill-name>/SKILL.md`". `--bare` skips skill auto-discovery.
- **Verdict:** `claude -p "<prompt>" --dangerously-skip-permissions --model <m>` is valid. Run it as a non-root user or in a recognized sandbox.

### 7b. Codex CLI. **Contradicted**: `--full-auto` no longer exists. Confidence: High

Latest npm release: 0.166.0, 2026-10-29. The official docs pages on developers.openai.com were blocked, so this comes from the source and the published binary.

- **Removed.** `--full-auto` is absent from `codex-rs/utils/cli/src/shared_options.rs` and `codex-rs/exec/src/cli.rs` at tag `rust-v0.166.0` and on main. Searching the strings of the 0.166.0 linux-x64 binary from npm (`@openai/codex@0.166.0-linux-x64`; not executed) finds 0 occurrences of `full-auto`.
- **History, by bisecting raw tags:**
  - Up to 0.125.0 (npm, 2026-04-24) it was a real flag: "Convenience alias for low-friction sandboxed automatic execution (--sandbox workspace-write)".
  - From 0.126.0 (the first npm release with it is 0.128.0, 2026-04-30) to 0.146.0 it was a hidden "Legacy compatibility trap for the removed `--full-auto` flag". That trap printed "warning: `--full-auto` is deprecated; use `--sandbox workspace-write` instead."
  - **From 0.147.0 (2026-08-07) it is gone entirely.** By clap's normal behaviour (not tested), passing it should fail as an unknown argument.
- **Still valid in 0.166.0:**
  - `codex exec [OPTIONS] [PROMPT]`
  - `--skip-git-repo-check` ("Allow running Codex outside a Git repository.")
  - `-m/--model`
  - `-s/--sandbox <mode>`
  - `--dangerously-bypass-approvals-and-sandbox`, alias `--yolo`
- In exec mode, `approval_policy` defaults to `Never` (`exec/src/lib.rs`).
- **AGENTS.md.** Verified in `codex-rs/core/src/agents_md.rs`: `DEFAULT_AGENTS_MD_FILENAME: &str = "AGENTS.md"` and override file `AGENTS.override.md`. It collects every `AGENTS.md` "from the project root down".
- **Replacement command:** `codex exec --sandbox workspace-write --skip-git-repo-check --model <m> "<prompt>"`. Inside an external container, Harbor uses `--dangerously-bypass-approvals-and-sandbox` instead.

### 7c. Gemini CLI. Verified: High, with a caveat

Latest npm release: 0.66.0, 2026-10-27. Source is `packages/cli/src/config/config.ts` at tag `v0.66.0`.

- `.option('yolo', {alias: 'y', ...'Automatically accept all actions (aka YOLO mode...'})`.
- `.option('prompt', {alias: 'p', ... 'Run in non-interactive (headless) mode with the given prompt.'})`.
- `.option('model', {alias: 'm'})`.
- `--yolo` and `--approval-mode` cannot be combined.
- **Caveat.** Lines 758-764: "Force approval mode to default if the folder is not trusted." So `--yolo` is silently downgraded in an untrusted folder. Use `--skip-trust`, which sets `GEMINI_CLI_TRUST_WORKSPACE=true`, or set that env var directly, as Harbor does.
- **GEMINI.md** (`docs/cli/gemini-md.md`): "Context files, which use the default name `GEMINI.md`". They are loaded from `~/.gemini/GEMINI.md` and from the workspace directories and their parents.
- **Verdict:** `gemini --yolo -p "<prompt>" --model <m>` is valid. Add `--skip-trust` for reliable headless runs.

---

## Not verified, or only partially verified

- **Debian trixie package pages:** exact current revision and the `libreoffice-calc-nogui` trixie page. packages.debian.org and the mirrors were blocked; the finding rests on search-index snippets.
- **LibreOffice release-note wiki text and official announcement dates** (wiki and libreoffice.org blocked). The function availability itself is verified from the help and core sources.
- **Rendered pages** on openpyxl.readthedocs.io, docs.harborframework.com, inspect.aisi.org.uk and developers.openai.com. The repository and package doc sources were used instead.
- **Codex runtime behaviour** of an unknown `--full-auto`. It was inferred from source and binary strings; the binary was not executed.
