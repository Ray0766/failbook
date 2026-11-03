# Failbook UI references: teardown, direction, visual system

> Research note from 2026-11-03. The screenshots it cites were taken during the session and are not published: three of the tools showed synthetic data made for the teardown, which must not be mistaken for real evaluation results. The site in `src/failbook/site.py` implements the direction and visual system proposed in sections 2 and 4. Checker messages and rule counts quoted in the mock-ups are as they were during the design pass; after the review fixes xlsx-001 has 7 rules and different messages (current text: `cases/*/failure/grade.json`).

Date viewed for every reference: **2026-11-03**.
Screenshots: `(session scratch, not published)/refs/`
Capture scripts: `(session scratch, not published)/scripts/`

Method: Playwright 1.63.0 Chromium, full-page captures at 1280px and 390px widths (`*-1280-*`, `*-390-*`). Very tall pages were also cut into segments or cropped to the top so they stay readable. Every visual claim below comes from a screenshot I opened. Computed font sizes come from `getComputedStyle` on the same pages (`scripts/typescale.js`).

---

## 0. What could and could not be reached

The session's egress policy blocked most of the requested hosts. A blocked CONNECT returned 403 for both curl and Chromium. I did not retry these hosts or route around the block, for example through mirrors, caches or archive copies. **I have not seen these pages and do not describe them.**

| Requested | Result | What I did instead |
|---|---|---|
| https://inspect.aisi.org.uk/log-viewer.html (docs) | Blocked (proxy 403) | Installed `inspect-ai==0.3.287` from PyPI (an allowed host). Ran a **synthetic** 3-sample eval on the built-in `mockllm` model and opened the **real Inspect log viewer** with `inspect view`. The docs page and any hosted demo were not seen. |
| https://www.swebench.com/ | Blocked (proxy 403) | None. Not seen. |
| https://programbench.com/ | Blocked (proxy 403) | None. Not seen. |
| https://incidentdatabase.ai/cite/1/ | Blocked (proxy 403) | None. Not seen, so the "case record" format below is my own proposal, not taken from AIID. |
| https://caniuse.com/css-grid | Blocked (proxy 403) | None. Not seen, so the history matrix below is my own proposal, not taken from caniuse. |
| Lighthouse viewer (googlechrome.github.io) | Blocked (proxy 403) | Installed `lighthouse@13.6.0` from npm. Ran it against the local Inspect viewer page and viewed the **real standalone HTML report** it writes. |
| GitHub Actions run page (github.com) | Blocked (github.com web returns 403) | Ran a **synthetic** 5-test Playwright suite and viewed the **real Playwright HTML test report** (1.63.0). |
| pypi.org project page (extra candidate) | Reachable, but Fastly served a CAPTCHA to headless Chromium. I did not try to solve it. | Skipped. Evidence: `blocked-pypi-captcha-1280.png` |
| https://jsr.io/@std/csv/score and /versions (extra candidate) | **Live page viewed.** I gave Chromium a scratch NSS trust store holding the session's proxy CA bundle. TLS verification stayed on. | n/a |

Caveat: the Inspect, Playwright and Lighthouse captures show real product UIs rendering **synthetic data I made** (xlsx-themed mock samples and tests). They show how each tool lays out evidence. They are not real evaluation results.

---

## 1. Teardown (specific part → principle → Failbook translation → where it does not apply)

### 1.1 Inspect log viewer (inspect-ai 0.3.287, local `inspect view`, synthetic log)

Screenshots: `inspect-loglist-1280-full.png`, `inspect-loglist-390-full.png`, `inspect-task-1280-full.png`, `inspect-task-390-full.png`, `inspect-sample-1280-full.png`, `inspect-sample-390-full.png`, `inspect-sample-scoring-1280-full.png`, `inspect-sample-scoring-390-full.png`, `focus-inspect-loglist-1280.png`

What I saw: a three-level drill-down (log list → task with samples grid → sample) in a dense data-grid UI. Text sizes are 9.6 to 14.4px; most text is 12.8px and IDs are monospace. The sample header pins Input and Answer side by side, with score pills such as `1 new-orders-added` and `0 totals-include-all-orders`. Tabs: Transcript / Messages / Scoring / Activity / Info / JSON. Prev/next arrows read "‹ Sample xlsx-001 (Epoch 1) ›". The Scoring tab is a table: Scorer | Answer | Score | Explanation.

| Specific part | Principle | Failbook translation | Where it does not apply |
|---|---|---|---|
| Sample header pins **Input** and **Answer** next to each other above every tab | Keep the claim and the task in view while reading the evidence | Case header shows the task (one line) and the agent's **claim** as a labelled quote ("Claim: agent output, not verified"), directly above the rule results | Not as a sticky header. On a document page it takes vertical space from evidence on phones. Show it once at the top, and show the claim again only inside the failing rule that contradicts it. |
| Scoring table puts the per-rule value next to an **explanation string** ("totals-include-all-orders: Total Qty is 198, expected 269 (difference 71)") | Every verdict carries its own human-readable evidence line | Each rule row: verdict, rule id, rule title, then the evidence sentence from `grade.json` `detail`, with cell references set in mono (`Orders!D15`) | Inspect puts all explanations for a sample into one text blob that is separate from the scores. Failbook keeps evidence **per rule**, in the same row as its verdict. |
| **JSON** tab next to the rendered views | Raw record one click away; the rendered view is a reading of the record, not a replacement for it | Link "View raw grade.json" beside each run, plus the case.yaml source link | No in-page JSON viewer. A plain link to the file in the repo is enough and works on static hosting. |
| Prev/next sample arrows in the header | Move sideways through siblings without going back to the list | "← xlsx-000 · xlsx-002 →" at the end of each case page | Fine as plain links. Keyboard shortcuts are not needed for 15 cases. |
| **Status column shows a green ✓ for every sample, including xlsx-001, which fails 2 rules** (`inspect-task-1280-full.png`) | Anti-pattern: "run completed" and "check passed" share one icon | Failbook keeps **run state** (completed / errored / not run) separate from **verdict** (pass / fail / inconclusive). A check mark only ever means "rule passed". | n/a |
| Scorer summary shows **mean 0.667 / 0.333** over 3 samples | Anti-pattern: three decimal places imply precision that n=3 cannot carry | Show counts: "2 of 3 runs fail". Never a rate or mean for small samples. | n/a |
| 390px: grid overflows (scrollWidth 450 at 390), the scorer box overlaps the Tags button, and the **Explanation column collapses to one character per line** (`inspect-sample-scoring-390-full.png`) | A multi-column evidence table has to reflow by stacking, not by shrinking columns | Below about 640px, every rule becomes a stacked block (verdict + title, then evidence at full width, then the expected/observed pairs) | Inspect is a desktop tool, so this is not a fault for its own audience. Failbook is a public reading site and must reflow at 320px. |

### 1.2 Playwright HTML test report (1.63.0, synthetic suite)

Screenshots: `pwreport-index-1280-full.png`, `pwreport-index-390-full.png`, `pwreport-filter-failed-1280-full.png`, `pwreport-filter-failed-390-full.png`, `pwreport-test-1280-full.png`, `pwreport-test-390-full.png`, `pwreport-flaky-1280-full.png`, `pwreport-flaky-390-full.png`, `focus-pwreport-1280.png`

What I saw: a search field plus a segmented filter bar with counts ("All 4 · ✓ Passed 1 · ✕ Failed 2 · ⚠ Flaky 1 · ⊘ Skipped 1"). Test rows put failures first; each row has a glyph, "describe › test" title, muted `file:line` and duration. The type scale is effectively two steps: 14px for everything (UI and mono errors) and a 32px test title. The test page has Run / Retry #1 tabs, each with its own glyph and duration. Below them: an Errors block (`Expected: 269` / `Received: 198`, mono, green and red), Test Steps (glyph, step name, `file:line`, duration bar), Attachments (my `cells.csv`), a "Copy prompt" button and previous/next links.

| Specific part | Principle | Failbook translation | Where it does not apply |
|---|---|---|---|
| Filter chips that carry **counts and a glyph**, plus free-text search, in one bar | Show the distribution while filtering; each chip states what it will reveal | Catalog toolbar: search input, then chips "Fails on latest retest (n) · Passes after fix (n) · Inconclusive (n)", then tag chips | The counts did not add up: All = 4 while 5 tests exist, because Skipped is left out of All. Failbook chip counts must add up to the total, or the bar must say what is excluded. |
| Filter state lives in the URL hash (`#?q=s:failed`); results line reads "Filtered: 2" | Filtered views can be shared and need no server | Catalog filter state in `?q=&tag=&verdict=` (or the hash), so it works on GitHub Pages. A live region says "Showing 3 of 15 cases". | Playwright's `s:` query mini-syntax is too obscure. Use visible chips and plain words. |
| **Run / Retry #1 tabs**: each attempt is kept with its own verdict; flaky shows ✕ then ✓ (`pwreport-flaky-1280-full.png`) | Attempts are appended, never overwritten; contradictory attempts stay visible | Retest history is an append-only list. A pass that follows a fail sits next to it, and nothing hides the first run. | Tabs hide all but one attempt. For history, a table that shows every run at once compares better. |
| `Expected: 269` / `Received: 198` in mono with labels | Expected vs observed are **labelled**, so the colour only adds to the meaning | Evidence pairs: "Expected 269 · Observed 198 · Difference −71" with text labels; colour is optional | Failbook uses "observed", not "received", and adds the **explanation** row that Playwright does not have |
| Attachment (`cells.csv`) listed under the failing test | Evidence files sit next to the result they support | The "Files" section lists the input workbook, the failure artifact and grade.json with size and SHA-256, and each failing rule links to the artifact | Inline previews of xlsx are out of scope. Use download links and cell excerpts. |
| "Copy prompt" button inside the error block | One-click copy for the next action | "Copy command" button on the re-run block; "Copied" is announced through `aria-live` | Do not copy an AI prompt. Copy the deterministic CLI command. |
| 390px: chips wrap to two lines (fine); the **error code frame is clipped** and needs horizontal scrolling; "next »" collides with the breadcrumb (`pwreport-test-390-full.png`) | Code and formula blocks need a wrap policy | Formulas and cell refs use `overflow-wrap:anywhere`. Shell commands scroll horizontally inside their own block, and only there. | n/a |
| Focus: the browser-default 1px blue outline on a chip (`focus-pwreport-1280.png`) | Visible, but thin | Use a stronger 2px outline with a 2px offset (§4.6) | n/a |

### 1.3 Lighthouse HTML report (Lighthouse 13.6.0 CLI, run against the local Inspect viewer)

Screenshots: `lighthouse-report-1280-full.png`, `lighthouse-report-390-full.png`, `lighthouse-1280-seg0.png`, `lighthouse-1280-seg1.png`, `lighthouse-1280-seg2.png`, `lighthouse-top-1280.png`, `lighthouse-top-390.png`, `lighthouse-audit-aria-required-children-1280.png`, `lighthouse-audit-aria-required-children-390.png`, `lighthouse-audit-errors-in-console-1280.png`, `lighthouse-audit-errors-in-console-390.png`, `focus-lighthouse-1280.png`

What I saw: a single self-contained HTML document. Body text is 14px, labels 12px and code 12 to 14px mono. A category gauge row sits on top. Each category section starts with its gauge and an honest caveat, then failing audits, then collapsed groups with counts: "Passed audits (24)", "Not applicable (36)", "Additional items to manually check (10)". Each audit row is a `<details>`. Opening one shows a one-sentence explanation, a "Learn more" link and an evidence table, for example "Failing Elements" with an element thumbnail and the DOM snippet in mono, or Source | Description for console errors.

| Specific part | Principle | Failbook translation | Where it does not apply |
|---|---|---|---|
| Legend "▲ 0–49 · ■ 50–89 · ● 90–100", with the **same shape and colour** reused on every audit and metric row; a hollow grey ○ marks informative audits | Status uses shape, colour and a legend, and the meaning holds in greyscale | Verdict glyphs with distinct shapes (§4.5), always followed by the word ("Fail"), with a short legend at the top of the rule list | No numeric score bands. A Failbook rule is pass / fail / inconclusive and has no 0–100 value. |
| Failing audits listed first; **passed, not-applicable and manual-check groups collapsed with counts** | The page spends its space on what is wrong; what passed is counted and one click away | Rule list: failing rules expanded first, then "Passed rules (4)" collapsed (`<details>`), then "Not checked (n)" | With only about 6 rules per case, the passed list can stay expanded on desktop. Collapse it only on narrow screens. |
| "**Additional items to manually check (10)**": automated and manual checks kept apart | Machine verdicts are separate from human judgement | A separate group "Human judgement (labelled, with reviewer and date)" under the automatic rule results; never mixed into the automatic count. This matches the AGENTS.md rule. | n/a |
| Caveat under the gauge: "Values are estimated and may vary"; under Accessibility: "Automatic detection can only detect a subset of issues" | State the limits of the measurement where the number is | Under the history table: "Curated cases are not an unbiased sample. n = 3 runs. Counts, not rates." | n/a |
| `<details>` audit → explanation → **evidence table** (element + snippet) | Summary row → why it matters → concrete evidence, all on one page with no navigation | Each failing rule expands to: what the rule checks (traced to a sentence of the task), evidence table Cell / Expected / Observed (`Orders!D15` · 269 · 198), and the formula found (`=SUM(F2:F11)`) | Element screenshot thumbnails do not apply. A small rendered cell excerpt (a 3–5 row slice of the sheet) is the equivalent, and it is optional. |
| One static HTML file that works offline | The report is itself the artifact | Each case page is fully static and readable without JS; JS only improves filtering and copy buttons | n/a |
| 390px: the 2-column metric grid becomes one column with values right-aligned; gauges wrap 3+2; evidence tables wrap inside the column (`lighthouse-top-390.png`, `lighthouse-audit-errors-in-console-390.png`) | A document layout reflows without horizontal scrolling | The same one-column collapse. Evidence tables with 3 short columns (cell, expected, observed) fit at 320px. | Animated score gauges and big coloured numbers are decorative stats for Failbook (§2); leave them out |

### 1.4 JSR package page: Score tab and Versions tab (live, https://jsr.io/@std/csv/score and https://jsr.io/@std/csv/versions)

Screenshots: `jsr-score-1280-full.png`, `jsr-score-1280-top.png`, `jsr-score-390-full.png`, `jsr-score-390-top.png`, `jsr-versions-1280-full.png`, `jsr-versions-1280-top.png`, `jsr-versions-390-full.png`, `jsr-versions-390-top.png`, `jsr-package-1280-full.png`, `jsr-package-1280-top.png`, `jsr-package-390-full.png`, `jsr-package-390-top.png`

What I saw: a rubric checklist. Each row has a ✓ or ✕ glyph, a criterion title, one sentence saying what is required (with a link), and a right-aligned **fraction** (`2/2`, `0/1`, `5/5`). A side card holds a 100% ring. The type scale is 14 / 16 / 18 / 30px in the DM Sans web font. Tabs show counts (Versions 49, Dependencies 1, Dependents 43). The Versions tab has a download chart above a list of one card per release, with relative dates ("Released by … 2 years ago") and the latest card tinted green.

| Specific part | Principle | Failbook translation | Where it does not apply |
|---|---|---|---|
| Rubric row = glyph + criterion + **what is required** + fraction | Each check says what it requires and shows partial results as counts | Rule row: glyph + word, rule title, "Checks: …" sentence, and for multi-cell rules a count such as "2 of 3 cells wrong" | n/a |
| **Header says "JSR Score 100%" and the ring says 100%, while "Has module docs in all entrypoints" shows ✕ 0/1** on the same page (`jsr-score-1280-full.png`). The page gives no reason. | Anti-pattern: a headline aggregate that contradicts a visible failure undermines trust in both | The case verdict is derived mechanically, and the rule that drives it is named: "Fail: 2 of 6 rules fail (totals-include-all-orders, totals-are-live)". No weighted score anywhere. | n/a |
| Tab labels with counts (Versions 49, Dependencies 1) | Show the size of each section before it is opened | Section headings with counts: "Retest history (3 runs)", "Files (3)" | Tabs hide content and are not needed. Use one scrolling page with anchors. |
| 390px: the score card moves above the checklist; tabs overflow into a "•••" menu; fractions stay right-aligned and legible (`jsr-score-390-full.png`) | The checklist pattern reflows cleanly | Rule rows keep the glyph on the left and the count on the right at 320px | Hiding navigation behind "•••" costs discoverability. Failbook has few sections, so it can wrap them. |
| Versions: **uniform cards**, relative dates ("2 years ago"), the latest release tinted green; at 390px the chart legend collides with the chart toolbar (`jsr-versions-390-top.png`) | Anti-pattern for an evidence record: a card wall wastes density, relative dates go stale on a static page, and green means "latest" here but "pass" elsewhere | Retest history is a compact table with absolute ISO dates (`2026-11-03`). It is ordered by date, never by result. Colour is used only for verdict. No chart. | n/a |
| Footer prints an absolute timestamp "It is 2026/11/03, 2:28:49 UTC" | Say when the page was produced | Footer: "Built 2026-11-03 from commit abc1234", which suits a static build | n/a |

---

## 2. Direction comparison and recommendation

| | A. Research archive (editorial case dossier; catalog as a dense table) | B. Debugging workbench (rule list left, evidence panel right, like a test runner) |
|---|---|---|
| Fits "arrive from a link, understand one failure" | Strong: reads top to bottom as claim → what is wrong → proof → reproduce → history | Medium: built for triaging many tests in one long session |
| 320px reflow | Natural (Lighthouse and JSR reflowed to one column cleanly) | Weak: both workbench references broke at 390 (Inspect overflow and one-character-per-line explanation; Playwright clipped code frame) |
| Static hosting, no JS, print, Ctrl+F, crawlable | Strong | Needs JS state for panel selection; panes with their own scroll break Ctrl+F and printing |
| Bilingual long text (task, why_missed) | Good: one column with a limited line length | Squeezed into a narrow pane |
| Scanning many rules | Good enough with fails first and passed collapsed with counts | Best |
| Risk of generic look | Can drift into a blog. Avoid that with a dense evidence table and mono cell refs. | Can drift into a SaaS dashboard |

**Recommendation: A, research archive, with workbench affordances built into the evidence section.** Most visits arrive from a link in an issue, paper or chat and need one case understood once, in order: what the agent claimed, what is actually wrong, the proof, and how to re-run it. A linear document matches that order, while a panel workbench suits repeated triage across hundreds of tests, which a 15-case catalog does not need. Both workbench-style references I viewed degraded at 390px. Inspect overflowed its grid and squeezed the explanation to one character per line, and Playwright clipped its error frame. A layout that depends on side-by-side panes works against the 320px reflow requirement. Lighthouse shows that a document can still be evidence-dense: failures first, passed and manual items collapsed with counts, an evidence table inside each audit, all in one static file that reflowed cleanly to one column. A static document also works with JS off, prints, responds to Ctrl+F and diffs in git, which suits GitHub Pages and an append-only record. From the workbench, Failbook keeps stable per-rule anchors, prev/next case links, glyph-plus-word verdicts and, at 1024px and wider, a sticky rule index in the left margin that jumps to evidence without becoming its own scroll pane. The catalog is where density pays off, so it should be a real `<table>` with filter chips that turns into stacked rows on narrow screens.

### 2.1 Page skeletons (proposal)

Catalog (`/`), 1280px:
```
Failbook                                  [search cases, rules, tags…]
Fails on latest retest (11)  Passes after fix (3)  Inconclusive (1)  | formulas  ranges  sort …
Showing 15 of 15 cases · curated, not an unbiased sample
ID        Case                                          Failing rules      Provenance     Latest retest
xlsx-001  New orders inserted above the Total row…      2 of 6             reconstructed  ✕ Fail · 2026-11-03
xlsx-002  Margin hardcoded …                            1 of 5             reconstructed  ✓ Pass · 2026-10-28
```
At less than 640px, each row becomes a block: ID + verdict on line 1, title on line 2, then "2 of 6 rules fail · reconstructed · 2026-11-03".

Case page (`/cases/xlsx-001/`):
```
xlsx-001 · New orders inserted above the Total row are left out of the total
✕ Fail on failure artifact: 2 of 6 rules fail      spreadsheet · easy · case v1 · reconstructed · failure kind: tool/model
[Claim]     "Done. I added the 3 new orders … The Total row now reflects all orders."
[Expected]  Total row sums all 13 orders, as formulas
[Observed]  =SUM(F2:F11) covers 10 orders; Total Qty 198
[Why]       insert_rows() moves cells but rewrites no formula …
Rule checks (6) — legend: ✕ fail  ✓ pass  ? inconclusive  – not run
 ✕ Fail  totals-include-all-orders   Total Qty is 198, expected 269 (difference 71)
          Cell        Expected        Observed
          Orders!D15  269             198
          Orders!F15  =SUM(F2:F14)    =SUM(F2:F11)
 ✕ Fail  totals-are-live             +10 on SO-1161 changed Total Qty by 0, expected +10
 ▸ Passed rules (4)
Human judgement (0) — none recorded
Re-run          $ failbook run xlsx-001 …              [Copy command]
Files (3)       input/sales.xlsx · failure/sales.xlsx · failure/grade.json — size · SHA-256 · download
Retest history (3 runs, newest first, append-only; counts not rates)
  Date        Agent / version     Model     Verdict   Rules failing   Record
← xlsx-000                                                       xlsx-002 →
```
(Synthetic layout. Values come from `cases/xlsx-001-total-misses-new-rows/failure/grade.json`; the history rows are placeholders, and the CLI flags still need to be confirmed against the real CLI.)

The history table is my own proposal. caniuse was not reachable, so I am not claiming a caniuse-style matrix. If one is wanted later, use rows = case, columns = agent+version and cells = "✕ 2/6 · 2026-11-03" (the date is visible text, not a tooltip). Never sort it by pass count. The header field "failure kind" is a placeholder: case.yaml does not yet record model / tool / environment / grader attribution, which AGENTS.md requires.

---

## 3. Patterns to avoid (seen in the references)

- One icon meaning both "completed" and "passed" (Inspect status ✓ on failing samples).
- Means and percentages on tiny n (Inspect "0.667" on 3 samples; JSR "100%").
- An aggregate that contradicts a visible failing item (JSR 100% next to ✕ 0/1).
- Filter counts that do not add up to "All" (Playwright All 4 vs 5 tests).
- Relative dates on records ("2 years ago") and card walls for history (JSR Versions).
- Uppercase, letter-spaced section labels (Inspect "SCORER", Lighthouse "DIAGNOSTICS"): they do not carry over to Chinese and should not be part of the type system.
- Sub-12px text for primary data (Inspect 9.6 to 11.2px).

---

## 4. Visual system (proposal)

### 4.1 Type scale (px, rem at a 16px root)

| Token | px / line-height | Use |
|---|---|---|
| `--fs-xs` | 13 / 1.45 | Table metadata, footnotes, file sizes. Never for verdicts or evidence. |
| `--fs-sm` | 14 / 1.5 | Catalog table cells, rule titles, evidence tables, mono code |
| `--fs-md` | 16 / 1.6 (zh 1.75) | Body: claim, expected/observed/explanation prose |
| `--fs-lg` | 20 / 1.35 | Section headings (Rule checks, Re-run, Files, History) |
| `--fs-xl` | 24 / 1.3 | Case title on narrow screens |
| `--fs-2xl` | 30 / 1.25 | Case title from 768px up; the only large size on the page |

Weights: 400 body, 600 headings and verdict words. No italics for zh text; use quotation marks 「」 or weight instead. Do not apply `text-transform` or letter-spacing to labels. Use `font-variant-numeric: tabular-nums` for counts, dates and cell values. Prose column max width is 72ch; tables may use the full content width, up to 1120px.

Font stacks (system fonts only, no CDN):
```css
--font-sans: system-ui, -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial,
             "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", "Noto Sans CJK SC",
             "Source Han Sans SC", "WenQuanYi Micro Hei", sans-serif,
             "Apple Color Emoji", "Segoe UI Emoji";
--font-mono: ui-monospace, "SFMono-Regular", "Cascadia Mono", "Segoe UI Mono", Menlo,
             Consolas, "Liberation Mono", "DejaVu Sans Mono",
             "PingFang SC", "Microsoft YaHei", "Noto Sans Mono CJK SC", monospace;
:lang(zh-Hant) { --font-sans: system-ui, -apple-system, "Segoe UI", Roboto, Arial,
             "PingFang TC", "Microsoft JhengHei", "Noto Sans CJK TC", sans-serif; }
```
Put `lang="en"` / `lang="zh-Hans"` on each bilingual block so line-height, quotes and font selection follow the language. Sheet names in cell references can be CJK (`'订单'!D15`), which is why the mono stack ends with CJK faces.

### 4.2 Spacing scale (4px base)

`--sp-1: 4px · --sp-2: 8px · --sp-3: 12px · --sp-4: 16px · --sp-5: 24px · --sp-6: 32px · --sp-7: 48px · --sp-8: 64px`

Page gutter is 16px under 640px and 24px from 640px up. Rule row padding is 12px vertical and 16px horizontal. Sections are 48px apart. Table cell padding is 8px × 12px. Dividers are 1px lines; there are no card shadows.

### 4.3 Colour tokens (WCAG contrast computed with the relative-luminance formula)

| Token | Light | Dark | Meaning | Contrast (light / dark) |
|---|---|---|---|---|
| `--bg` | `#FAFAF7` | `#141413` | Page | n/a |
| `--surface` | `#FFFFFF` | `#1C1C1A` | Tables, rule rows | n/a |
| `--sunken` | `#F1F0EB` | `#252522` | Code, command and formula wells | n/a |
| `--text` | `#1B1B18` | `#ECEAE3` | Body | 16.5 / 15.3 on bg |
| `--muted` | `#595851` | `#A8A69E` | Metadata | 6.8 / 7.6 on bg |
| `--border` | `#D6D4CC` | `#3A3934` | Decorative dividers only | n/a |
| `--border-strong` | `#85837B` | `#7A786F` | Input outlines, chip outlines (1.4.11 needs ≥ 3:1) | 3.8 / 3.9 on surface |
| `--link` | `#1D4ED8` | `#8AB4F8` | Links (always underlined in prose) | 6.4 / 8.8 on bg |
| `--focus` | `#1D4ED8` | `#8AB4F8` | Focus ring | ≥ 5.8 / ≥ 6.5 on every surface and tint |
| `--fail` / `--fail-tint` | `#B3261E` / `#FCEBE9` | `#F2938A` / `#3A1E1B` | Rule or verdict fails | 6.3 on bg, 5.7 on tint / 8.2, 6.7 |
| `--pass` / `--pass-tint` | `#1E6B3C` / `#E6F2EA` | `#7FCF98` / `#1B3223` | Rule or verdict passes | 6.2, 5.7 / 9.9, 7.4 |
| `--inconclusive` / `-tint` | `#8A5300` / `#FBF0DB` | `#E9B95A` / `#382C14` | Check could not decide (grader error, timeout, ambiguous output) | 6.1, 5.6 / 10.1, 7.5 |
| `--neutral` / `-tint` | `#595851` / `#EDECE7` | `#A8A69E` / `#2A2A27` | Not run, not applicable, superseded | 6.8, 6.0 / 7.6, 5.9 |

Every verdict foreground passes 4.5:1 (1.4.3 AA) on bg, surface, sunken and its own tint, in both themes. Body text on any tint is at least 11.3:1. Colour is used **only** for verdicts, links and focus, never for decoration and never to mean "latest". Failure *attribution* (model / tool / environment / grader) is a text field, not a colour. Dark mode follows `prefers-color-scheme`, with a manual `data-theme` override.

### 4.4 Density

The catalog is one row per case at about 44px tall, which fits about 15 cases on a 1280×800 screen. The case page shows the verdict line and the claim in the first screen and the first failing rule's evidence by the second. Evidence values and formulas are mono at 14px; nothing that carries meaning is below 13px.

### 4.5 Verdicts without relying on colour (1.4.1)

| Verdict | Glyph (inline SVG, `aria-hidden`) | Visible word (en / zh) | Container |
|---|---|---|---|
| Fail | ✕ in a square | Fail / 不通过 | `--fail` text, `--fail-tint` background, 1px `--fail` left border |
| Pass | ✓ in a circle | Pass / 通过 | `--pass` text, no tint (passes should not draw attention) |
| Inconclusive | ? in a diamond | Inconclusive / 无法判定 | `--inconclusive` text, tint, dashed left border |
| Not run | – (en dash) in a hollow circle | Not run / 未运行 | `--neutral` text |

- The word is always visible, so the glyph is redundant and the colour is a third channel. The shapes still differ in greyscale. The fail and pass foregrounds have almost the same luminance on purpose: there is no hidden "lighter = better" reading.
- Screen readers read the text: `<span class="verdict verdict--fail"><svg aria-hidden="true">…</svg>Fail</span>`. Rules use `<li>` inside `<ol aria-label="Rule checks">`, with the verdict word first.
- Expected vs observed is shown by labels ("Expected", "Observed", "Difference"), never by red/green text alone. A wrong value also carries a text marker such as "≠ expected".
- Counts are written as "2 of 6 rules fail" and "3 runs: 2 fail, 1 pass". No percentages, bars or rankings. The history is sorted by date only and carries the caveat "n = 3, curated cases; not a reliability estimate".
- A `:target` rule (deep link `#rule-totals-include-all-orders`) gets a 2px `--focus` outline and a "Linked rule" text badge, not only a background tint.

### 4.6 Focus (2.4.7 Focus Visible, 2.4.11 Focus Not Obscured (Minimum))

- `:focus-visible { outline: 2px solid var(--focus); outline-offset: 2px; }` on every interactive element, including `<summary>` and the copy buttons. Never `outline: none` without a replacement. The ring clears 3:1 against every surface and tint (§4.3).
- If a header or rule index is sticky, set `html { scroll-padding-top: <header height + 8px>; }` and keep sticky regions out of the way of the focused element: the sticky rule index sits beside the content, not on top of it. On narrow screens nothing is sticky.
- Add a skip link "Skip to rule checks". The tab order follows the reading order. Filter chips are `<button aria-pressed>`.

### 4.7 Reflow at 320 CSS px (1.4.10)

- Test at 320px and at 1280px with 400% zoom. Nothing may scroll horizontally at page level (Inspect reached scrollWidth 450 at a 390 viewport).
- Catalog `<table>` becomes stacked row blocks below 640px. The table stays in the DOM, with `display:block` rows and a visible label for each value. The evidence table (Cell / Expected / Observed) keeps 3 columns at 320px; when a formula is long, the row stacks into labelled lines.
- Mono formulas and cell refs use `overflow-wrap:anywhere`. Only shell command blocks may scroll horizontally, inside their own focusable `<pre tabindex="0">`.
- Long CJK titles wrap normally; `word-break: keep-all` is never used on zh text.

### 4.8 Motion (prefers-reduced-motion)

- Motion explains state only: a ≤150ms opacity change when a `<details>` opens, and a ≤150ms colour change on hover and press. There are no entrance animations, animated gauges, count-up numbers, skeleton shimmer or smooth-scroll hijacking.
- "Copied" feedback is a text change announced by `aria-live="polite"`, with no animated toast.
- `@media (prefers-reduced-motion: reduce) { *, *::before, *::after { transition: none !important; animation: none !important; scroll-behavior: auto !important; } }`

---

## 5. Licensing and reuse note

Only structure and method are borrowed. No code, icons, fonts (for example JSR's DM Sans) or copy are reused. Failbook glyphs should be drawn fresh as inline SVG. The Inspect, Playwright and Lighthouse screenshots show tools rendering synthetic data generated for this teardown. They are research notes and should not be published as examples of real agent results.

## 6. Reproduction of the captures

- Inspect: `scripts/inspect_task/xlsx_mock.py` (mockllm, 3 synthetic samples, explicit token usage set because tokenizer download is blocked) → `tools/inspect/.venv/bin/inspect view start --log-dir tools/inspect/logs --port 7575` → `scripts/inspect_nav.js`, `scripts/shot.js`
- Playwright report: `tools/pwreport/playwright.config.js`, `tools/pwreport/tests/xlsx.spec.js` (synthetic) → `scripts/pw_nav.js`
- Lighthouse: `npx lighthouse http://127.0.0.1:7575/ --output html` with `CHROME_PATH` set to Playwright's Chromium → `scripts/segments.js`, `scripts/lh_detail.js`
- JSR: `HOME=tools/chromehome node scripts/pypi_jsr.js` and `scripts/jsr_versions.js` (scratch NSS store with the session CA bundle; TLS verification stays on)
- Focus and type data: `scripts/focus.js`, `scripts/typescale.js`
(All paths are relative to `(session scratch, not published)/`.)
