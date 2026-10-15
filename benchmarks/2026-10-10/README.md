# Benchmark: docx-cli vs Anthropic's docx skill vs OfficeCLI (2026-10-10)

Three tools, one model, six real document tasks, three runs each. This page has the full results and everything you'd want to know before trusting them.

## Setup

- **Model:** Claude Haiku 5.5 drives every run (the cheapest current tier, where tool ergonomics matter most).
- **Tools:**
  - docx-cli 0.27.0
  - Anthropic's docx skill, staged from `anthropics/skills` at commit `dbd4588f9e1033efb41dad4bef2f7947c8993d44`
  - OfficeCLI 1.0.156, binary SHA-256 `8a6c1aa383ee8b9069d012914418a4814d183727ca100b6c3d1ef2839a460290`
- **Tasks** (one starting `.docx` each, identical across tools), each written as a human would delegate it, with no tool vocabulary:
  - **mnda:** fill an NDA's cover page, keeping its fonts.
  - **invoice:** fill an invoice table, restructure it, and replace the logo.
  - **resume:** restyle a résumé (headings, spacing, margins) and fill it in.
  - **contract-markup:** redline a contract with tracked changes and comment on clauses.
  - **contract-finalize:** accept some tracked changes, reject the rest, reply to and resolve comments.
  - **eliot-journal:** author a poetry journal from scratch, with columns, verse, footnotes, links, and a figure.
- **Grading:**
  - Microsoft Word renders every output to page images.
  - An Opus judge grades each one against a written rubric for that task, using the renders as the primary evidence plus inspection of the file.
  - A task counts as **solved** only on a full `success` verdict; `partial` counts as not solved.
- **Metrics** come from each run's transcript. "Input tokens (cache-weighted)" is `fresh + 1.25 × cache-write + 0.1 × cache-read`: the cost-relevant input, in normal-input-token equivalents. Raw counts are below.

## Results

| | **docx-cli 0.27.0** | **Anthropic docx skill** | **OfficeCLI 1.0.156** |
| :-- | --: | --: | --: |
| Tasks solved (of 6) | 6.0 (6.0–6.0) | 6.0 (6.0–6.0) | 4.7 (4.0–5.0) |
| Rendered correctly (of 6) | 6.0 (6.0–6.0) | 6.0 (6.0–6.0) | 6.0 (6.0–6.0) |
| Effective input tokens | 902k (899k–904k) | 1.15M (1.09M–1.22M) | 1.92M (1.78M–2.00M) |
| Output tokens | 32k (27k–35k) | 54k (48k–66k) | 63k (59k–66k) |
| Wall-clock (sum of 6 tasks) | 914 s (808 s–990 s) | 1,254 s (1,039 s–1,511 s) | 1,834 s (1,659 s–2,085 s) |
| Tool calls | 98 (98–98) | 124 (118–128) | 165 (157–174) |

| Scenario | docx-cli 0.27.0 pass@1 · pass@3 · pass^3 | Anthropic docx skill pass@1 · pass@3 · pass^3 | OfficeCLI 1.0.156 pass@1 · pass@3 · pass^3 |
| :-- | :--: | :--: | :--: |
| mnda | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ |
| invoice | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ | 2/3 · ✓ · ✗ |
| resume | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ | 0/3 · ✗ · ✗ |
| contract-markup | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ |
| contract-finalize | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ |
| eliot-journal | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ |
| **all 6** | 18/18 · 6/6 · 6/6 | 18/18 · 6/6 · 6/6 | 14/18 · 5/6 · 4/6 |

| Run | Solved | Partial | Rendered | Layout sound | Eff. input | Output | Wall-clock | Tool calls | Breaches/relay |
| :-- | --: | --: | --: | --: | --: | --: | --: | --: | :-- |
| docx-cli r1 | 6/6 | 0 | 6/6 | 6/6 | 904k | 27k | 808 s | 98 | 0/no |
| docx-cli r2 | 6/6 | 0 | 6/6 | 6/6 | 899k | 35k | 943 s | 98 | 0/no |
| docx-cli r3 | 6/6 | 0 | 6/6 | 6/6 | 902k | 33k | 990 s | 98 | 0/no |
| anthropic-docx-skill r1 | 6/6 | 0 | 6/6 | 6/6 | 1.14M | 48k | 1,039 s | 125 | 0/no |
| anthropic-docx-skill r2 | 6/6 | 0 | 6/6 | 6/6 | 1.09M | 48k | 1,214 s | 118 | 0/no |
| anthropic-docx-skill r3 | 6/6 | 0 | 6/6 | 6/6 | 1.22M | 66k | 1,511 s | 128 | 0/no |
| officecli r1 | 4/6 | 2 | 6/6 | 5/6 | 2.00M | 66k | 1,756 s | 174 | 0/no |
| officecli r2 | 5/6 | 1 | 6/6 | 5/6 | 1.78M | 62k | 1,659 s | 157 | 0/no |
| officecli r3 | 5/6 | 1 | 6/6 | 5/6 | 1.99M | 59k | 2,085 s | 165 | 0/no |

All figures are the mean across 3 runs, summed over the 6 tasks; ranges are min–max across runs.

**Per task** (mean of 3 runs: cache-weighted input tokens / wall-clock / tool calls):

| Task | docx-cli | Anthropic docx skill | OfficeCLI |
| :-- | --: | --: | --: |
| mnda (fill) | **142k** / **120 s** / **15** | 159k / 164 s / 17 | 292k / 353 s / 25 |
| invoice (fill + restructure) | **191k** / **214 s** / **19** | 273k / 330 s / 26 | 470k / 417 s / 36 |
| resume (restyle + fill) | **153k** / **175 s** / **15** | 238k / 271 s / 23 | 304k / 298 s / 28 |
| contract-markup (redline + comment) | **154k** / **156 s** / **16** | 202k / 229 s / 22 | 253k / 204 s / 22 |
| contract-finalize (accept/reject + reply) | **94k** / **86 s** / **10** | 179k / 159 s / 21 | 178k / 129 s / 18 |
| eliot-journal (author from scratch) | 167k / 162 s / 22 | **98k** / **102 s** / **15** | 426k / 432 s / 36 |

On the five tasks that edit an existing document, docx-cli used about 30% fewer input tokens and 35% less time than the Anthropic skill. On the one from-scratch authoring task, the skill was cheaper on every measure. All three tools passed both contract tasks in every run.

**Raw input token counts** (mean per run, all 6 tasks):

| | docx-cli | Anthropic docx skill | OfficeCLI |
| :-- | --: | --: | --: |
| Fresh input | 183 | 232 | 324 |
| Cache write | 381k | 424k | 597k |
| Cache read | 4.25M | 6.18M | 11.76M |

docx-cli's three runs all totaled 98 tool calls by coincidence. The per-task counts differ (r1: 15, 17, 19, 15, 11, 21; r2: 18, 20, 11, 18, 11, 20; r3: 13, 20, 16, 15, 9, 25).

## What failed

OfficeCLI accounted for every miss:

- **résumé, 3 of 3 runs:**
  - The structural work was right (data filled, headings unified through the style, margins set).
  - But after the margin change every location/date line wrapped, some mid-word ("Cambrid / ge, MA"). The task asks for a clean one-page résumé with right-aligned dates.
  - This is a known trap: tab-aligned text that overflows after a margin change. docx-cli detects it and repairs it when margins change, so this task favors a tool that has met it before.
- **invoice, 1 of 3 runs:** the new notes sentence went into an empty spacer paragraph, and the old notes line was left below it.

## Caveats

- **We built the tool and the test.** docx-cli's author wrote the six tasks and the rubrics. These tasks are also docx-cli's regression suite: features are developed against them. Expect home-field advantage. The cleanest evidence here is the comparison between the two competitors, and the cost gap on tasks all three tools solved.
- **The judge was not blind.** It was told which tool produced each result. Competitor runs carried an explicit note: grade from the Word renders, treat docx-cli-specific inspection commands in the rubric as optional, and don't penalize a different internal representation of the same result.
- **The prompts differ by tool.** The docx-cli and Anthropic-skill prompts mirror each other: same role, same rules, same report, each pointing at its own tool. The docx-cli prompt also names its orientation commands (`--help`, `info locators`) and says to prefer batch operations. OfficeCLI got a shorter prompt that points it at its own SKILL.md and help.
- **Staging the Anthropic skill:**
  - Upstream `anthropics/skills` no longer ships the `unpack.py`/`pack.py` helpers, and the skill as staged used `unzip`/`zip` directly.
  - Claude Code's runtime runs Python with `-I`, which breaks the skill's sibling-module imports.
  - The skill still solved every run. Its token and time figures may include some overhead from working around this.
- **Small sample.** One model, three runs per tool, one machine (macOS, Word as the renderer), one day.
- **Correctness converged on current models.** Earlier runs on older models showed a large correctness gap: in July 2026 (Haiku 4.5, docx-cli v0.20.4, published in the README at the time) docx-cli solved 5.0/6 and the default skill 0.7/6. That run's raw data was not published. On Haiku 5.5 both tools solve everything; what's left is cost, speed, and consistency.

## Raw data

Everything the tables above are computed from is in this folder:

- `aggregate.json` / `aggregate.md`: per-tool summaries and the per-run table.
- `runs/<tool>-r<N>/<task>/`, one per run and task:
  - `verdict.json`: the judge's structured verdict.
  - `review.md`: the judge's written review, with the evidence it checked.
  - `metrics.json`: tokens, wall-clock, and tool calls, taken from the transcript.

The full agent transcripts and rendered pages aren't published: they're large and contain machine-local paths.

## Reproducing

The harness lives in [`.claude/skills/weak-agent-test`](../../.claude/skills/weak-agent-test): the scenarios, the rubrics (`criteria.md`, withheld from the agents during runs), the staging scripts for each competitor, and the metrics and aggregation scripts. Running it needs Claude Code, Microsoft Word for the renders, and API access for Haiku (exercise) and Opus (judge).
