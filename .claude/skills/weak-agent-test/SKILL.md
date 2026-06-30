---
name: weak-agent-test
description: "Run the weak-agent (Haiku) adversarial test harness against docx-cli. Spawns Haiku agents to perform real document tasks over six scenarios — five editing (MNDA form-fill + font fidelity, invoice table-edit/restructure + logo replace, résumé styling, contract redlining + commenting, contract finalize via accept/reject + comment reply/resolve) and one authoring (T. S. Eliot poetry journal: multi-column, verse, footnotes, links, figure) — renders every result with Word, judges them, measures the Haiku tool economy (docx-cli vs other calls), and synthesizes a prioritized ergonomics report. Use when the user says 'adversarial review', 'test docx-cli with weak agents', 'run the haiku harness', 'weak agent test', or wants to re-run yesterday's adversarial process."
allowed-tools: Bash, Read, Write, Glob, Workflow
---

# Adversarial review — weak-agent harness for docx-cli

This harness answers one question: **can weak agents (Haiku) actually use docx-cli to
get real work done, and what should we fix first?** It runs the `weak-agent-test`
workflow (`.claude/workflows/weak-agent-test.js`), which fans out one Haiku agent
per scenario, renders every output with Microsoft Word, grades each against
ground-truth criteria, and synthesizes a prioritized improvement report.

The test corpus is **bundled with this skill** under `scenarios/`, one folder per
scenario, named after its key (`scenarios/mnda/`, `scenarios/loi/`, …). Each scenario
folder is self-describing and holds everything that scenario needs:

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
(key, bucket label, `edit`/`author` kind, the doc filename, whether to render a
baseline); the actual request/criteria/fixture/assets all live in the folder. The skill is
therefore self-contained and travels with its test corpus. To change what a scenario
tests, edit the files in its folder. (Heavy, ephemeral run outputs — edited docx,
renders, reviews, the report — are dumped to `./tmp/docx-weak-agent-test/<ts>/`,
never into the repo.)

Each run produces, under the timestamped run dir, **one result folder per scenario**
(named after its key) plus the run-level report and metrics:

```
<RUN_DIR>/
  REPORT.md            ← synthesized report (+ appended measured metrics)
  haiku-metrics.md     ← measured per-Haiku-agent tokens/time/tool split
  haiku-metrics.json
  benchmark.json       ← self-reported tool economy
  <key>/               ← one per scenario; the worked-on copy lives here
    task.md  assets/   ← (criteria.md is withheld from this copy — judge-only)
    <doc>.docx         ← the edited/authored document
    renders/output/    ← the Word-rendered PNGs the judge looked at
    renders/baseline/  ← pristine before-render (baseline scenarios only, e.g. psa)
    review.md          ← the judge's saved review for this task (written in-run)
    metrics.json       ← this task's measured tokens/time/tool split (written post-run by haiku-metrics)
```

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

Create an empty timestamped `./tmp/` run dir. **Do NOT copy the scenarios here** — the
workflow's **Stage** phase copies _only the active scenarios' folders_ from the pristine
`$SCENARIOS_DIR` into `$RUN_DIR`, one subfolder per scenario (`$RUN_DIR/<key>/`), so
originals stay untouched, the repo stays clean, and a single-scenario run doesn't drag
the whole corpus along:

```bash
TS="$(date +%Y.%m.%d-%H%M%S)"
RUN_DIR="./tmp/docx-weak-agent-test/$TS"
mkdir -p "$RUN_DIR"   # empty; the workflow's Stage phase seeds one subfolder per active scenario from $SCENARIOS_DIR
echo "RUN_DIR=$RUN_DIR"
```

### 3. Launch the workflow

Invoke the `Workflow` tool with `scriptPath` pointing at the workflow file and pass
the absolute paths as `args`:

```
Workflow({
  scriptPath: "<REPO>/.claude/workflows/weak-agent-test.js",
  args: {
    runDir: "<RUN_DIR from step 2>",
    binary: "<BINARY from step 1>",
    scenariosDir: "<SCENARIOS_DIR from step 1>",
    only: <optional scenario filter — see below>
  }
})
```

`only` restricts the run to a subset of scenarios (omit it to run all 6). To run a
**single task**, pass its key as a plain string — `only: "mnda"`. It also accepts an
array (`only: ["mnda", "loi"]`) or a comma/space-separated string (`only: "mnda,loi"`);
all three are normalized to the same list, so use whichever is convenient. The keys are
the folder names under `$SCENARIOS_DIR` (run `ls "$SCENARIOS_DIR"` if you need to
confirm them); unknown keys abort the run with a "No scenarios matched" error listing
the valid ones.

> Use `scriptPath`, NOT `name: "weak-agent-test"`. Launching by name resolves to a
> copy cached at session start, so any edit to the workflow made during the session is
> ignored; `scriptPath` always reads the current file from disk. (The workflow also
> tolerates `args` arriving as a JSON string — the runtime stringifies it — so passing
> a plain object is fine.)

When the tool returns, **note the `Transcript dir:` path it prints** — call it
`TRANSCRIPT_DIR` (it looks like `…/subagents/workflows/wf_<id>`). You need it in
step 4 to measure per-agent tokens and time.

**Scenario keys** (omit `only` to run all 6):
`mnda`, `invoice`, `resume`, `contract-markup`, `contract-finalize`, `eliot-journal`.

If the user passed scenario keys as arguments to this skill (e.g.
`/weak-agent-test mnda loi` or `mnda,loi`), parse them into the `only` array.
Otherwise run everything.

The run is heavy (6 Haiku agents, serial Word rendering, 6 judges + a synthesis pass);
it can take many minutes. Watch live progress with `/workflows`.

### 4. Save the report + measure the Haiku metrics

When the workflow completes, its return value is
`{ report, runDir, binary, exercises, verdicts, benchmark }`. The `report` already
contains the scoreboard, per-task merits/demerits, and the self-reported tool economy.
Two more steps add the _measured_ numbers:

1. Write `report` (Markdown) to `"$RUN_DIR/REPORT.md"` and `benchmark` (JSON) to
   `"$RUN_DIR/benchmark.json"` with the Write tool.
2. Run the metrics pass over the run's transcripts to append the **accurate
   per-Haiku-agent tokens + wall-clock time + docx/non-docx tool split** (the workflow
   can't measure these itself — the runtime gives the script no token API and bans
   clocks, so this reconstructs them from the transcripts). `TRANSCRIPT_DIR` is the
   path you noted in step 3:
   ```bash
   bun "$REPO/.claude/skills/weak-agent-test/scripts/haiku-metrics.ts" \
     "<TRANSCRIPT_DIR>" "$RUN_DIR" "$BINARY" "haiku" >> "$RUN_DIR/REPORT.md"
   ```
   The 4th argument is the exercise model the run used (the workflow's `args.model`,
   default `haiku`). **If you ran the workflow with a non-default `args.model` (e.g.
   `sonnet`), pass that same value as the 4th arg** — otherwise the metrics default to
   `haiku`, match no exercise agents, and emit an empty table (the script warns on stderr
   when 0 transcripts match the filter).
   This appends the measured "Haiku tool & cost economy" table to REPORT.md, writes
   `$RUN_DIR/haiku-metrics.json` (run-level), and drops each scenario's measured row
   into `$RUN_DIR/<key>/metrics.json` so every per-task folder is self-contained.
   Token cost is reported as **effective input**
   (cache-weighted: fresh input + cache write ×1.25 + cache read ×0.1) plus **output**,
   kept separate — NOT a single "total tokens", because cache reads are ~10× cheaper
   than fresh input and lumping them in overstates cost. The raw cache split is in the
   Totals table and `haiku-metrics.json`.
3. Present in chat: the **Executive summary**, the **per-task merits/demerits**, and
   the **Haiku metrics** (total docx vs other calls + docx share, effective input +
   output tokens, total time, and the per-scenario outliers). Tell the user where the
   artifacts live:
   - `<RUN_DIR>/REPORT.md` — findings + scoreboard + per-task merits/demerits + measured metrics table
   - `<RUN_DIR>/haiku-metrics.json` and `benchmark.json` — the raw numbers
   - `<RUN_DIR>/<key>/` — one folder per scenario, each holding that task's worked-on
     `.docx`, its `renders/` (the Word PNGs the judge looked at), `review.md` (the
     judge's saved review), and `metrics.json` (that task's measured tokens/time/tool split)

## Notes

- This harness is **re-runnable**: each invocation rebuilds the binary (mandatory),
  stages a fresh `./tmp/` run dir, and never mutates the bundled `scenarios/`.
- The headline **benchmark metric** is the Haiku tool economy — docx-cli calls vs
  other tool calls (self-reported per agent, aggregated by the workflow). A high
  non-docx share, or a lot of docx calls for a simple task, is a friction signal.
- The weak agents invoke the binary at an allowlisted absolute path
  (`dist/docx`), so they should not hit permission prompts for the CLI itself. The
  benign shell commands they and the render step use (`mkdir`, `cp`, `ls`, `cat`) are
  NOT yet allowlisted — if you get prompted, add them via the `update-config` skill or
  run with edits allowed. See `.claude/settings.local.json`.
- To add a scenario, create a folder under this skill's `scenarios/<key>/` holding
  `task.md` (the agent-facing request, in human voice — NO tool vocabulary, so the
  agent must discover the features), `criteria.md` (the judge-only grading rubric —
  withheld from the agent's run workspace, read by the judge from the pristine source),
  the fixture `.docx` (edit scenarios only), and an `assets/` folder, then add a
  routing entry to `SCENARIOS` in the workflow (`.claude/workflows/weak-agent-test.js`):
  `{ key, bucket, kind, doc, baseline }`. To change what an existing scenario tests,
  edit the files in its folder — the request/criteria/fixture/assets all live there,
  not in the workflow.
