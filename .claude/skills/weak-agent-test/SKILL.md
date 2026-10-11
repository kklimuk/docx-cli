---
name: weak-agent-test
description: "Run the weak-agent adversarial test harness against docx-cli. Spawns weak exercise agents (Haiku by default, Sonnet to probe, or a local agent harness's pre-produced runs) to perform real document tasks over six scenarios — five editing (MNDA form-fill + font fidelity, invoice table-edit/restructure + logo replace, résumé styling, contract redlining + commenting, contract finalize via accept/reject + comment reply/resolve) and one authoring (T. S. Eliot poetry journal: multi-column, verse, footnotes, links, figure) — renders every result with Word, has opus judge them against ground-truth rubrics, measures each exercise's tool economy, token cost, wall-clock, and correctness (from transcripts for Claude, the exercise.json ledger for the local harness), and synthesizes a prioritized ergonomics report. Use when the user says 'adversarial review', 'test docx-cli with weak agents', 'run the haiku harness', 'weak agent test', or wants to re-run yesterday's adversarial process."
allowed-tools: Bash, Read, Write, Glob, Workflow
metadata:
  internal: true
---

# Adversarial review — weak-agent harness for docx-cli

This harness answers one question: **can weak agents actually use docx-cli to get
real work done, and what should we fix first?** It runs the `weak-agent-test`
workflow (`.claude/workflows/weak-agent-test.js`), which fans out one weak exercise
agent per scenario (Haiku by default — swappable to Sonnet via `args.model`), renders
every output with Microsoft Word, grades each against ground-truth criteria with an
**opus** judge, and has **opus** synthesize a prioritized improvement report.
Exercise agents do NOT self-report tool counts — every tool-economy and token number
is **measured** after the run (agents under-count their own calls ~2×, so self-reports
were dropped): from the agent transcripts for the Claude arms, from each scenario's
`exercise.json` ledger for the local arm. Both roll up into the same Run-metrics table
(tokens, wall-clock, tool split, correctness) via `exercise-metrics.ts`.

The test corpus is **bundled with this skill** under `scenarios/`, one folder per
scenario, named after its key (`scenarios/mnda/`, `scenarios/invoice/`, …). Each
scenario folder is self-describing and holds everything that scenario needs:

- `task.md` — the AGENT-FACING request, written as a human delegating the work:
  the goal, the data, the intent — and **no tool vocabulary** (no `docx` commands,
  locators, or OOXML terms), because discovering which features deliver the outcome
  is part of what's measured,
- `criteria.md` — the JUDGE-ONLY grading rubric (the precise, tool-specific checks).
  The stage step **withholds it from the agent's run workspace**, and the judge reads
  it from the pristine source — the agent never sees the answer key,
- the fixture `.docx` to work on (edit scenarios only; authoring scenarios create
  their output fresh),
- `assets/` — any additional inputs (data files, images; empty for most edit
  scenarios).

The workflow's `SCENARIOS` manifest holds only the per-scenario **routing** metadata
(key, bucket label, `edit`/`author` kind, the doc filename); whether a baseline gets
rendered is DERIVED from the kind (every `edit` scenario has a pristine source, so it
gets one — see `hasBaseline()`), not a stored field. The actual
request/criteria/fixture/assets all live in the folder. The skill is
therefore self-contained and travels with its test corpus. To change what a scenario
tests, edit the files in its folder. (Heavy, ephemeral run outputs — edited docx,
renders, reviews, the report — are dumped to `./tmp/docx-weak-agent-test/<ts>/`,
never into the repo.)

Staging is ONE code path for every backend: `scripts/stage-scenario.ts` copies a
scenario folder, strips the judge-only `criteria.md`, and verifies the inputs landed.
The workflow's Stage agent runs it per scenario; the local corpus runner imports it.

Each run produces, under the timestamped run dir, **one result folder per scenario**
(named after its key) plus the run-level report and metrics:

```
<RUN_DIR>/
  REPORT.md            ← synthesized report + the measured run-metrics section, both
                          written by your post-run exercise-metrics.ts --append-report
  exercise-metrics.md  ← measured per-exercise-agent tokens/time/tool split
  exercise-metrics.json
  <key>/               ← one per scenario; the worked-on copy lives here
    task.md  assets/   ← (criteria.md is withheld from this copy — judge-only)
    scratch/           ← the agent's PRIVATE scratch dir (batch JSONL, drafts). The
                          session scratchpad is shared by every concurrent agent, so
                          same-named files there collide — see "Isolation breaches"
    <doc>.docx         ← the edited/authored document
    renders/output/    ← the OUTPUT: Word-rendered page PNGs + read.md (markdown read view)
    renders/baseline/  ← the pristine "before": page PNGs + read.md (every EDIT scenario;
                          absent only for the authored eliot-journal — no source to diff)
    review.md          ← the judge's saved review for this task (written in-run)
    verdict.json       ← the judge's structured verdict incl. taskSuccess (written in-run
                          by the judge — the correctness source the metrics pass reads)
    metrics.json       ← this task's measured tokens/time/tool split + correctness
                          (written by your post-run metrics pass)
```

The **render step fires the moment each task finishes** (for both arms) and produces,
for the OUTPUT and — whenever a pristine source exists (every edit scenario) — its
BASELINE "before", BOTH deliverables in each render dir: the page PNGs AND a `read.md`
(the markdown read view of that doc). The judge reads all four (output PNGs + read.md,
baseline PNGs + read.md) to compare before/after both visually and textually. The
workflow's render step is **idempotent**: for the **local backend** the corpus runner
already produced the SAME artifacts at the SAME paths as it went, so the render step
just reuses them (re-rendering only anything missing) — no double-render; for the
**Claude backend** nothing is pre-rendered, so it does the full Word render. Either
way the judge grades Word-rendered PNGs (the local harness runs on the mac, where the
corpus's default render engine IS Word).

## Steps

Run these in order from the repo root. Do NOT skip the build — the global `docx` on
PATH is a stale binary; the harness must test the CURRENT working tree.

### 1. Preflight — ALWAYS rebuild (mandatory gate)

The whole harness is meaningless if it tests a stale binary, so the build is a hard
gate, not an optional step. **Always run `bun run build:binary`, even if `dist/docx`
already exists** — never reuse a prior build. Abort the whole run if any check below
fails.

```bash
REPO="$(git rev-parse --show-toplevel)"
cd "$REPO"
SCENARIOS_DIR="$REPO/.claude/skills/weak-agent-test/scenarios"   # this skill's bundled corpus (one folder per scenario)

# Word must be installed (this harness renders with Word, not LibreOffice).
test -d "/Applications/Microsoft Word.app" || echo "WARNING: Microsoft Word not found — render phase will fail."

# (1) Build the CURRENT working tree into a fresh standalone binary. Abort on failure.
bun run build:binary || { echo "BUILD FAILED — abort"; exit 1; }
BINARY="$REPO/dist/docx"

# (2) Hard gate: the fresh binary must match package.json's version AND have `render`.
# `--version` prints "docx X.Y.Z"; take the 2nd space-delimited field. NOTE: use `cut`,
# NOT an awk field reference — a literal dollar-N positional token gets clobbered by
# slash-command positional-arg substitution when this skill runs with arguments, mangling
# the gate. Keep this whole block free of dollar-N tokens for the same reason.
EXPECTED="$(bun -e 'console.log(require("./package.json").version)')"
GOT="$("$BINARY" --version | cut -d' ' -f2)"
echo "built docx $GOT (package.json: $EXPECTED)"
[ "$GOT" = "$EXPECTED" ] || { echo "VERSION MISMATCH ($GOT != $EXPECTED) — build is stale, abort"; exit 1; }
"$BINARY" render --help >/dev/null 2>&1 || { echo "render MISSING — build stale/broken, abort"; exit 1; }
echo "preflight OK: fresh $GOT binary with render"
```

If the version mismatches or `render` is missing, the build did not reflect the
working tree — stop and fix it before running. Do not proceed on a stale binary.

> First-run note: Word-for-Mac rendering triggers a one-time macOS **Automation**
> permission prompt for the controlling terminal. If the render phase fails on a
> fresh machine, grant it under System Settings → Privacy & Security → Automation and
> re-run.

### 2. Make an isolated run workspace (under ./tmp/)

Create an empty timestamped `./tmp/` run dir **per workflow run**. **Do NOT copy the
scenarios here** — the workflow's **Stage** phase runs `scripts/stage-scenario.ts` for
_only the active scenarios_, seeding one subfolder per scenario (`$RUN_DIR/<key>/`), so
originals stay untouched, the repo stays clean, and a single-scenario run doesn't drag
the whole corpus along:

```bash
TS="$(date +%Y.%m.%d-%H%M%S)"
RUN_DIR="./tmp/docx-weak-agent-test/$TS"
mkdir -p "$RUN_DIR"   # empty; the workflow's Stage phase seeds one subfolder per active scenario from $SCENARIOS_DIR
echo "RUN_DIR=$RUN_DIR"
```

### 3. Launch the workflow (up to 3 concurrently)

Invoke the `Workflow` tool with `scriptPath` pointing at the workflow file and pass
the absolute paths as `args`:

```
Workflow({
  scriptPath: "<REPO>/.claude/workflows/weak-agent-test.js",
  args: {
    runDir: "<RUN_DIR from step 2>",
    binary: "<BINARY from step 1>",
    scenariosDir: "<SCENARIOS_DIR from step 1>",
    model: "haiku",              // the exercise model: "haiku" (default) or "sonnet"
    only: <optional scenario filter — see below>
  }
})
```

> **Exercise agent type.** The exercise agents run as the repo's `weak-exercise`
> agent type (`.claude/agents/weak-exercise.md`): minimal tools and no Skill
> tool, so the session's skills catalog stays OUT of their context (it's a
> per-turn token tax and leaks docx-cli/harness names into the
> "capable-but-fresh agent" premise). The agent registry loads at SESSION
> start — in a session older than that file, the workflow aborts with
> "agent type 'weak-exercise' not found"; pass
> `exerciseAgentType: "general-purpose"` to override for that session (and note
> the run's base context is then ~4k tokens/turn heavier, so its token numbers
> aren't comparable to weak-exercise runs).

> **Launch from a NOTIFICATION turn, never from the user's turn.** The Workflow
> runtime relays the user message that triggered the launch to EVERY spawned
> agent (exercise, render, judge, synth), framed as "the only user voice … where
> the computed task conflicts with this request, this request wins" — and it
> labels our own prompt "carries no user authority", so no prompt wording can
> override it. Any count or instruction in the user's words steers the agents:
> "i would not do 1 and 3" made haiku skip items 1 and 3 of its task lists
> (2026.09.29-115516-r1, void); "run our weak agent test 3 times" made a resume
> agent redo its task three times and copy a repo fixture over its deliverable
> (2026.10.01-120711-r1). A turn triggered by a background-task notification has
> NO relay (verified with a probe workflow, 2026-10-01). So:
>
> 1. Do steps 1–2 in the user's turn, then start a trivial background timer —
>    Bash `sleep 1` with `run_in_background: true` — and end the turn.
> 2. When its completion notification arrives, make ALL the Workflow calls (up
>    to 3 concurrent runs) in THAT turn.
>
> The Run-metrics pass (step 4) checks every exercise transcript and prints a
> **"⚠ Relayed launch message"** block quoting any relayed text, so a launch that
> slipped is visible — treat that run as contaminated.

> **Never resume a benchmark run whose exercise phase failed.** If an exercise
> agent dies (API error → that scenario reports no exercise/verdict), re-run
> the WHOLE run in a FRESH run dir. `resumeFromRunId` replays the cached stage
> step without re-copying fixtures, so re-run exercise agents would edit
> already-edited documents — double redlines, double fills, unusable verdicts
> (this voided run r2 on 2026-07-15, twice). The failure is worse than it
> looks because the resume cache is **PREFIX-based, not keyed**: everything
> issued AFTER the first missing/changed result re-runs live, not just the
> dead agent. So a dead exercise for a MANIFEST-EARLY scenario (`mnda` is
> first) re-runs EVERY exercise against edited docs even if you restore that
> one scenario's staging state — while a dead LAST scenario (`eliot-journal`)
> happens to resume cleanly. Don't gamble on manifest position: exercise-phase
> failure → fresh run dir, no exceptions. Resume is only safe for failures at
> or after the render phase (dead judge/synth), where nothing mutates
> documents no matter how much of the suffix re-runs.

**Running 3 at a time (the fast path to averaged numbers).** The benchmark
methodology is 3 runs per arm/model, and runs can go **concurrently**: launch up to
three Workflow invocations in one message, each with its OWN `RUN_DIR` from step 2
(suffix the timestamp, e.g. `$TS-r1`, `$TS-r2`, `$TS-r3`). This is safe because the
only shared mutable resource is Microsoft Word, and the CLI itself serializes Word
access across processes with an advisory lock (`src/core/render/engines/word-mac.ts`)
— concurrent runs' renders queue instead of corrupting each other. Don't go beyond ~3:
renders start spending more time queueing than rendering. A haiku-vs-sonnet
comparison is just two batches: three runs with `model: "haiku"`, three with
`model: "sonnet"` (never mix models within one run dir).

`only` restricts the run to a subset of scenarios (omit it to run all 6). To run a
**single task**, pass its key as a plain string — `only: "mnda"`. It also accepts an
array (`only: ["mnda", "invoice"]`) or a comma/space-separated string; all forms are
normalized to the same list. The keys are the folder names under `$SCENARIOS_DIR`
(run `ls "$SCENARIOS_DIR"` if you need to confirm them); unknown keys abort the run
with a "No scenarios matched" error listing the valid ones.

> Use `scriptPath`, NOT `name: "weak-agent-test"`. Launching by name resolves to a
> copy cached at session start, so any edit to the workflow made during the session is
> ignored; `scriptPath` always reads the current file from disk. (The workflow also
> tolerates `args` arriving as a JSON string — the runtime stringifies it — so passing
> a plain object is fine.)

(Launch from the timer's notification turn — see the "Launch from a NOTIFICATION turn" box above.) When the tool returns, **note each run's `Transcript dir:` path it prints** — call it
`TRANSCRIPT_DIR` (it looks like `…/subagents/workflows/wf_<id>`). You need it in
step 4 to measure per-agent tokens and time. With concurrent runs, keep each
run's `(RUN_DIR, TRANSCRIPT_DIR)` pair matched.

**Scenario keys** (omit `only` to run all 6):
`mnda`, `invoice`, `resume`, `contract-markup`, `contract-finalize`, `eliot-journal`.

If the user passed scenario keys as arguments to this skill (e.g.
`/weak-agent-test mnda invoice`), parse them into the `only` array. Otherwise run
everything.

Each run is heavy (6 exercise agents, serialized Word rendering, 6 opus judges + an
opus synthesis pass); it can take many minutes. Watch live progress with `/workflows`.

### 4. Save the report + measure the exercise metrics

When a workflow completes, its return value is
`{ arm, report, runDir, binary, exercises, verdicts }`. The `report` contains the
scoreboard, per-task merits/demerits, and prioritized fixes — deliberately **without**
tool-call or token numbers (nothing self-reports them).

**One post-run command writes `REPORT.md` and the metrics — for BOTH backends.**
The judge already wrote each `<key>/verdict.json` in-run. The synth agent does NOT
write `REPORT.md`: Claude Code refuses report-file writes from subagents ("Subagents
should return findings as text"), so the synth only RETURNS its report, and the
workflow journal (`<TRANSCRIPT_DIR>/journal.jsonl`) records it. The workflow's JS can't
measure tokens/time either (no token API, no clocks), and `TRANSCRIPT_DIR` is only
known after launch — so both land in one script pass you run now:

1. **Run `exercise-metrics.ts --append-report`** per run (match each RUN_DIR with its
   own TRANSCRIPT_DIR). It writes `REPORT.md` = the synth's report (pulled from the
   journal, any apology preamble stripped) + the measured **Run metrics** section. It's
   idempotent — re-running replaces the metrics section rather than stacking another.
   - **Claude** (`exerciseBackend: "claude"`) — the 4th arg is the exercise model
     (`args.model`, default `haiku` — **pass `sonnet` if you ran sonnet**, or it matches
     no agents and emits an empty table):
     ```bash
     bun "$REPO/.claude/skills/weak-agent-test/scripts/exercise-metrics.ts" \
       "<TRANSCRIPT_DIR>" "$RUN_DIR" "$BINARY" "haiku" --append-report
     ```
   - **Local** (`exerciseBackend: "local"`) — numbers come from each `<key>/exercise.json`
     `_local` block; `--journal` points it at the synth's report:
     ```bash
     bun "$REPO/.claude/skills/weak-agent-test/scripts/exercise-metrics.ts" \
       --local "$RUN_DIR" "<harness/model label>" --journal "<TRANSCRIPT_DIR>" --append-report
     ```
   If the journal has no synth result, `REPORT.md` gets the metrics section only.
   If the workflow DID return a `report`, paste it above that section by hand (a
   re-run keeps it). If the synth died, there is no `report` either: resume the run
   (safe for a dead synth — see step 3) and re-run this command with
   `--journal <the resumed run's TRANSCRIPT_DIR>`.

   **Check the "⚠ Isolation breaches" block** the metrics section adds when any
   exercise agent touched the shared session scratchpad or `/tmp`, another
   scenario's folder, or any other repo path (`repo: tests/fixtures/…`, the pristine
   `scenarios/` corpus, older runs under `tmp/`). A **COLLISION** (the same shared scratchpad or `/tmp` path used by two scenarios) means one
   agent may have applied the other's batch — run 2026.09.30-101057-r1's mnda ran
   resume's `edits.jsonl` and then rewrote whole cells trying to recover. Check the
   transcript before reading that scenario's result as tool behavior. (Agents are
   told to use `<key>/scratch/`; the check catches the ones that don't.)
2. **Metrics** detail — the **measured per-exercise tokens (input AND output) +
   wall-clock + docx/non-docx tool split + correctness**.
   Either way you end up with the **Run metrics** section on `REPORT.md`, run-level
   `$RUN_DIR/exercise-metrics.{md,json}` (tagged with `backend`), and each scenario's
   measured row in `$RUN_DIR/<key>/metrics.json`. Token cost is reported as **effective
   input** (cache-weighted: fresh/non-cache input + cache write ×1.25 + cache read
   ×0.1) plus **output**, kept separate — NOT a single "total tokens", because cache
   reads are ~10× cheaper than fresh input and lumping them in overstates cost. The raw
   cache split is in the Totals table and `exercise-metrics.json`.
3. Present in chat: the **Executive summary**, the **per-task merits/demerits**, and
   the **measured metrics** — correctness (N/6 success), total docx vs other calls +
   docx share, fresh/cache input + output tokens, total wall-clock, and the
   per-scenario outliers. For a multi-run batch, also give the across-runs averages
   (tasks solved of 6, effective input, output, wall-clock). Tell the user where the
   artifacts live:
   - `<RUN_DIR>/REPORT.md` — findings + scoreboard + per-task merits/demerits + measured metrics table
   - `<RUN_DIR>/exercise-metrics.json` — the raw numbers
   - `<RUN_DIR>/<key>/` — one folder per scenario, each holding that task's worked-on
     `.docx`, its `read.md` (markdown read view) and `renders/` (the Word PNGs the
     judge looked at), `review.md` + `verdict.json` (the judge's saved review + verdict),
     and `metrics.json` (that task's measured tokens/time/tool split + correctness)

## Backends & arms

The exercise slot is **swappable**; everything downstream (render → opus judge →
opus synthesis, all against the same rubrics) is identical for every backend and
arm — that's what makes the numbers comparable.

- **Exercise model** (`args.model`): `"haiku"` (default) or `"sonnet"` — same
  workflow, same prompts, only the exercise agents' model changes.
- **Local harness** (`args.exerciseBackend: "local"`): the exercises run OUT OF BAND
  on the local-first agent harness (model built in), then the workflow
  renders/judges/synthesizes the results identically. Two steps:
  1. `bun "$REPO/.claude/skills/weak-agent-test/scripts/run-local-corpus.ts" "$SCENARIOS_DIR" "$RUN_DIR" "$BINARY" <HARNESS_DIR> [--context N] [--timeout SEC] [key...]`
     — serial (single GPU); stages via the same `stage-scenario.ts`, runs the
     harness per scenario, and parses each session ledger into
     `$RUN_DIR/<key>/exercise.json` (it also writes a run-level `$RUN_DIR/corpus.log`
     orchestration log itself — no stdout redirect needed). Every number is
     ledger-MEASURED (the local model
     is never asked to self-report), including a code-computed `status`
     (`completed` = the harness process ran to its own stop, `failed` = the watchdog
     killed it or it crashed on a signal — lifecycle only; the judge owns quality).
     `LOCAL_MODEL_PATH`/`LOCAL_MMPROJ_PATH` env vars override the harness's built-in
     model for control runs.
  2. Collect the results DETERMINISTICALLY and pass them to the workflow inline, so it
     skips its LLM LOAD agent and the code-computed `status`/account reach the judge
     straight from disk:
     ```bash
     EXERCISES="$(bun "$REPO/.claude/skills/weak-agent-test/scripts/collect-exercises.ts" "$RUN_DIR")"
     ```
     then launch with `{ runDir, binary, scenariosDir, exerciseBackend: "local",
     modelLabel: "<harness/model name>", exercises: <the collected array> }`. The
     workflow runs the normal Render/Judge/Synthesize pipeline on them. (If you omit
     `exercises`, the workflow falls back to an LLM LOAD agent that reads the
     exercise.json files itself — the `status` is still code-computed on disk, but
     prefer the deterministic collect so nothing re-reads it through a model.)
  The point of this arm is marketing the local harness by its **competitiveness with
  Haiku**: same tasks, same judge, same rubrics — only the exercise brain differs.
  Its cost/effort is ledger-measured into each `exercise.json` under `_local`, and the
  post-run `exercise-metrics.ts --local … --journal <TRANSCRIPT_DIR> --append-report`
  (step 4) rolls it up into the SAME Run-metrics table the Claude arms get — tokens,
  wall-clock, tool split, correctness — so the local-vs-Haiku numbers are directly
  comparable.
- **Competitor arm** (`args.arm: "anthropic-docx-skill"`): the A/B bake-off against
  Anthropic's bundled docx skill. First provision it with
  `bun "$REPO/.claude/skills/weak-agent-test/scripts/stage-competitor.ts" <SKILL_DEST> [RUN_DIR]` (fetches the real skill and
  installs/verifies its full toolset — fairness gate), then pass
  `arm: "anthropic-docx-skill", competitorDir: "<SKILL_DEST>"`. Only the exercise
  agents' tool instructions differ; grading is identical.

## Notes

- This harness is **re-runnable**: each invocation rebuilds the binary (mandatory),
  stages a fresh `./tmp/` run dir, and never mutates the bundled `scenarios/`.
- The headline **benchmark metrics** are correctness (tasks solved of 6), the tool
  economy (docx-cli calls vs other calls), and token cost as **effective input +
  output** — all measured by `exercise-metrics.ts` (transcripts for Claude, the
  `_local` ledger for local), never self-reported.
- The weak agents invoke the binary at an allowlisted absolute path
  (`dist/docx`), so they should not hit permission prompts for the CLI itself. The
  benign shell commands they and the render step use (`mkdir`, `cp`, `ls`, `cat`,
  `bun`) are NOT yet allowlisted — if you get prompted, add them via the
  `update-config` skill or run with edits allowed. See `.claude/settings.local.json`.
- To add a scenario, create a folder under this skill's `scenarios/<key>/` holding
  `task.md` (the agent-facing request, in human voice — NO tool vocabulary, so the
  agent must discover the features), `criteria.md` (the judge-only grading rubric —
  withheld from the agent's run workspace, read by the judge from the pristine source),
  the fixture `.docx` (edit scenarios only), and an `assets/` folder, then add a
  routing entry to `SCENARIOS` in the workflow (`.claude/workflows/weak-agent-test.js`,
  shape `{ key, bucket, kind, doc }`) AND to the `MANIFEST` in
  `scripts/run-local-corpus.ts` (shape `{ key, doc, kind }`). To change what an existing
  scenario tests, edit the files in its folder — the request/criteria/fixture/assets all
  live there, not in the workflow.

## Scripts

All Bun/TypeScript (this is a Bun-first repo — no shell scripts):

- `scripts/stage-scenario.ts` — stage ONE scenario (copy + strip criteria.md +
  verify). The single staging path: the workflow's Stage agent runs it; the local
  corpus runner imports it.
- `scripts/exercise-metrics.ts` — post-run run-metrics rollup (both backends):
  measured tokens (fresh/cache input + output), wall-clock, docx/other tool split, and
  correctness (from the judge verdicts), per scenario + totals + run-over-run
  comparison. Claude reads the transcripts; `--local <runDir> <label>` reads each
  `exercise.json` `_local` block.
- `scripts/run-local-corpus.ts` — run the exercise phase on the local agent harness
  (serial, watchdogged), producing `exercise.json` per scenario (with a code-computed
  `status`) for `exerciseBackend: "local"`.
- `scripts/collect-exercises.ts` — deterministically read the run's `exercise.json`
  files into the `args.exercises` array, so the workflow's local backend skips its
  LLM LOAD agent (the `status` reaches the judge from disk, not via a model).
- `scripts/parse-local-ledger.ts` — parse one local-harness session ledger into the
  exercise shape (ledger-measured tool calls, tokens, timings, and the process
  `status`).
- `scripts/local-exercise-prompt.md` — the prompt template the local runner renders
  per scenario (task inlined for the small model).
- `scripts/stage-competitor.ts` — provision the Anthropic docx skill + its full
  toolset for the competitor arm (fairness gate).
