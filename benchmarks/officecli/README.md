# Word task comparison

The latest comparison covers docx-cli 0.27.0, the Anthropic docx skill, and OfficeCLI 1.0.156, with three runs per tool. See the [October 10 results and limitations](results/2026-10-10.md) and [per-attempt data](results/2026-10-10.json). The [October 8 comparison](results/2026-10-08.md) and [earlier Anthropic-skill experiments](../../guides/benchmarks.md) remain separate historical records.

## What this measures

Can the same agent complete the same Word task using each tool? The six existing scenarios cover NDA filling, invoice editing, résumé styling, contract redlining, review finalization, and document authoring. Each has a task, pristine input/assets, and a separate judge-only rubric under [the existing corpus](../../.claude/skills/weak-agent-test/scenarios).

This corpus has already informed docx-cli development. Publish it as a comparison on our regression tasks, not an independent general-purpose benchmark. Before making broad superiority claims, add a separately sourced holdout corpus, freeze its rubrics before running either tool, and report it separately.

## Harness checks

`bun run test:benchmark` checks prompt isolation, required OfficeCLI setup, the save-before-render instruction, existing-arm compatibility, exercise-only metrics selection, and TypeScript types without downloading anything or calling a model. CI runs these checks.

## Setup

Requirements: Bun, Git, network access for initial staging, Microsoft Word for rendering, and a Claude environment with the `Workflow` tool and the existing `weak-exercise` agent. Read the [harness instructions](../../.claude/skills/weak-agent-test/SKILL.md). This workflow cannot be executed by running the JavaScript file directly with Bun; `agent`, `phase`, and `log` come from the Workflow runtime.

Check `command -v claude` and `claude --version` before starting. If multiple Claude installations exist, use the explicit path to the runtime that provides `Workflow` and background completion notifications. A coordinator waiting on a timer is not evidence that any benchmark exercise has started.

From the repository root:

```sh
# Pin an explicit release. Staging never installs globally or uses a moving latest.
bun benchmarks/officecli/stage.ts v1.0.156 tmp/officecli-v1.0.156

# Rebuild and snapshot docx-cli; record hashes, worktree diff, corpus, and versions.
# Both destinations must be new. Parent directories must already exist.
bun benchmarks/officecli/prepare.ts tmp/officecli-v1.0.156 tmp/officecli-comparison
```

The stage script verifies the downloaded binary against GitHub's release SHA-256 digest, fetches the matching source/skills at the release commit, checks its version, and records `manifest.json`. The prepare script checks that the binary and main skill still match, rebuilds docx-cli, snapshots the corpus, scripts, workflow and exercise-agent definition, and writes six run configurations: three repetitions per arm. Retain the dirty-worktree patch; disclose it with results rather than labeling a development build as the published release.

These setup commands cover the two CLI arms. The October 10 experiment also used the harness's `anthropic-docx-skill` arm, staged with [stage-competitor.ts](../../.claude/skills/weak-agent-test/scripts/stage-competitor.ts) and three additional fresh run directories. That arm receives the complete skill and its dependencies through `competitorDir`. Preserve its source revision and dependency versions for future runs; the October 10 staging did not retain the upstream commit, as disclosed in its report.

## Run

Use the Workflow runtime to execute each configuration, in alternating order: docx-cli r1, OfficeCLI r1, OfficeCLI r2, docx-cli r2, docx-cli r3, OfficeCLI r3. Read a configuration file as JSON and pass that object as `args`:

```js
Workflow({
  scriptPath: "<RUN_ROOT>/harness/weak-agent-test.js",
  args: /* contents of <RUN_ROOT>/officecli-r1.json */
})
```

**Launch from a background-task notification turn**, following the harness skill. The Workflow runtime otherwise relays the coordinator's launch request into exercise agents and contaminates their task context. Start a short background timer in the coordinator's user turn, end that turn, and launch workflows only when its completion notification arrives. Audit every exercise transcript for a relayed launch message and cross-scenario scratch access; invalidate contaminated runs.

Use the frozen exercise definition in `harness/agents.json` (Claude CLI: `--agents <RUN_ROOT>/harness/agents.json`). Use the same exercise model, judge model, resource limits, and runtime configuration for every arm. Save the exact resolved model IDs from transcripts; aliases can change. Start new sessions with the same base context. The exercise agent must not see grading rubrics or other arms' outputs. The harness stages private copies with `criteria.md` removed.

OfficeCLI receives its release-matched skill, specialized skills, built-in help, and access to linked documentation. It may use its own raw-XML commands, but neither arm may substitute another tool's editing library. Auto-updates are disabled with `OFFICECLI_SKIP_UPDATE=1`; normal resident behavior remains enabled. The exercise and render instructions require `close FILE` before Word reads the output, avoiding false failures from unflushed changes.

Word renders both outputs and pristine baselines. Both arms use the same judge and rubrics. Commands named in a rubric are inspection suggestions, not required implementations. A docx-cli reader failure must not count as an OfficeCLI failure. For structural facts such as pending revisions, inspect the OOXML independently when the reader is inconclusive. The current judge is not blinded to the arm; disclose that limitation and use blinded human review before a headline claim.

If an exercise fails due to authentication, API, or infrastructure errors, retain the failed run as infrastructure evidence and restart with fresh input copies. Never resume an exercise on already-edited inputs or count an unrun task as a tool-quality failure.

## Measure and report

Before each batch and before publishing results, check the frozen inputs:

```sh
bun benchmarks/officecli/verify.ts "<RUN_ROOT>"
```

A hash mismatch invalidates the comparison; do not mix tasks or tool builds across repetitions.

For each completed workflow, retain its transcript directory. The existing collector takes the executable under test as its third argument; use the OfficeCLI executable for the OfficeCLI arm:

```sh
bun "<RUN_ROOT>/harness/skill/scripts/exercise-metrics.ts" \
  "<TRANSCRIPT_DIR>" "<ARM_RUN_DIR>" "<EXECUTABLE_UNDER_TEST>" haiku --append-report
```

The collector selects `exercise:<scenario>` agents by transcript metadata before filtering by model; staging and rendering can use the same model and must not be included. Expect exactly six distinct exercise rows per full run. Its legacy `docxToolCalls` label means Bash calls mentioning the supplied binary in an OfficeCLI run; it is not a count of individual shell commands, and aliases can evade this heuristic. Prefer total tool calls for the headline comparison. Check transcript classification before publishing the totals. Fresh input, cache-write input, cache-read input, and output tokens should remain separate. The collector's historical weighted-input calculation is a proxy, not a dollar bill; use recorded provider cost or dated model pricing for dollar comparisons.

Publish per-task outcomes and all three runs, including failures, alongside aggregate means and ranges. Report:

- Correct requested edits, unchanged surrounding content, and formatting preservation.
- Word opening/rendering, pending revisions, comments, and accept/reject behavior where required.
- Measured exercise latency, tool calls, token usage, retries, and provider cost if available.
- Setup/version manifest, task/input hashes, prompts, outputs, renders, verdicts, and redacted transcripts.

Do not infer a population-wide win or statistical significance from six tasks. Do not silently modify a rubric, repair an output, or tune one arm after observing results; start a new experiment and label it if changes are needed.
