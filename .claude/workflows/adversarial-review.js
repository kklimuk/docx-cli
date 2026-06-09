export const meta = {
	name: "adversarial-review",
	description:
		"Weak-agent (Haiku) adversarial test of docx-cli: 6 scenarios over real fixtures + an authoring task, rendered with Word, judged, and synthesized into a prioritized ergonomics report.",
	phases: [
		{
			title: "Stage",
			detail:
				"copy each active scenario's folder (task.md, brief.md, fixture, assets/) into its own run-dir subfolder",
		},
		{
			title: "Exercise",
			detail: "one Haiku agent per scenario performs its task with docx-cli",
			model: "haiku",
		},
		{
			title: "Render",
			detail:
				"render each output to PNG via Word as its exercise finishes — serialized through a 1-slot gate (single Word instance)",
		},
		{
			title: "Judge",
			detail:
				"a strong agent grades each render + the Haiku transcript, writes a review.md into the scenario's result folder, firing as soon as that render is done",
		},
		{
			title: "Synthesize",
			detail: "one strong agent writes the prioritized improvement report",
		},
	],
};

// ---------------------------------------------------------------------------
// args, injected by the skill: { runDir, binary, scenariosDir, only? }
//   runDir        — absolute /tmp dir; the workflow stages one subfolder PER
//                   active scenario into it (<runDir>/<key>/…) and that subfolder
//                   doubles as the scenario's result folder.
//   binary        — absolute path to the freshly-built dist/docx
//   scenariosDir  — absolute path to the PRISTINE bundled scenarios dir. Each
//                   scenario is a folder named after its key, containing task.md
//                   (task + resolution criteria), brief.md (detailed instructions),
//                   the fixture .docx (edit scenarios only), and assets/ (extra
//                   inputs). Also the source of render baselines.
//   only          — optional scenario filter: run just these key(s). Accepts an
//                   array (["mnda","loi"]), a single key ("mnda" — the natural way
//                   to run ONE task), or a comma/space-separated string ("mnda,loi").
// ---------------------------------------------------------------------------
// The runtime delivers `args` as a JSON STRING (not the parsed object the tool
// docs imply), so parse it if needed. Accept an already-parsed object too.
const parsedArgs = typeof args === "string" ? JSON.parse(args) : args || {};
const { runDir, binary, scenariosDir } = parsedArgs;
const only = normalizeOnly(parsedArgs.only);
if (!runDir || !binary || !scenariosDir) {
	throw new Error(
		"adversarial-review requires args { runDir, binary, scenariosDir }",
	);
}

// Orchestration manifest, keyed by scenario folder name. The CONTENT of each
// scenario — its task, resolution criteria, brief, fixture, and extra assets —
// lives in `<scenariosDir>/<key>/` (task.md, brief.md, the .docx, assets/), NOT
// here. This manifest holds only what the workflow needs to ROUTE each scenario:
//   key       — folder name; also the result-folder name under runDir.
//   bucket    — human label for the scenario's category (prompt headers, scoreboard).
//   kind      — "edit" (work a staged copy of `doc` in place) | "author" (create `doc` fresh).
//   doc       — the .docx filename inside the scenario folder. For edit scenarios
//               it is the pristine fixture (staged + edited in place AND, when
//               baseline, rendered as the before). For author scenarios it does
//               NOT exist in the scenario folder — the agent creates it in the run dir.
//   baseline  — render the pristine `doc` as a before/after comparison (edit only).
const SCENARIOS = [
	{
		key: "mnda",
		bucket: "Form filling + highlight removal + font fidelity",
		kind: "edit",
		doc: "mnda.docx",
		// Baseline: this is the in-place-preservation fidelity test, so render the
		// pristine doc as a "before" and diff it against the filled "after".
		baseline: true,
	},
	{
		key: "invoice",
		bucket: "Table editing + restructure + image replace",
		kind: "edit",
		doc: "invoice.docx",
		baseline: false,
	},
	{
		key: "resume",
		bucket: "Styling fidelity + drawing preservation",
		kind: "edit",
		doc: "resume.docx",
		baseline: false,
	},
	{
		key: "contract-markup",
		bucket: "Legal review: redlining + commenting",
		kind: "edit",
		doc: "contract.docx",
		baseline: false,
	},
	{
		key: "contract-finalize",
		bucket: "Legal review: accept/reject + resolve comments",
		kind: "edit",
		doc: "contract-redlined.docx",
		baseline: false,
	},
	{
		key: "eliot-journal",
		bucket: "Authoring: columns, verse, footnotes, links, figure",
		kind: "author",
		doc: "journal.docx",
		baseline: false,
	},
];

const active =
	only && only.length
		? SCENARIOS.filter((scenario) => only.includes(scenario.key))
		: SCENARIOS;

if (!active.length) {
	throw new Error(
		`No scenarios matched ${JSON.stringify(only)}. Known keys: ${SCENARIOS.map((scenario) => scenario.key).join(", ")}`,
	);
}

log(
	`Adversarial review: ${active.length} scenario(s) — ${active.map((scenario) => scenario.key).join(", ")}`,
);
log(`Binary under test: ${binary}`);
log(`Run workspace: ${runDir}`);

// Schemas are defined here (not at the bottom) because `const` is not hoisted —
// they're referenced by the agent() calls below, so they must exist first. The
// prompt builders ARE at the bottom: function declarations hoist, so newspaper
// ordering still holds for them.
const STAGE_SCHEMA = {
	type: "object",
	additionalProperties: false,
	required: ["staged", "missing"],
	properties: {
		staged: {
			type: "array",
			description: "Scenario result folders that now exist after copying.",
			items: { type: "string" },
		},
		missing: {
			type: "array",
			description:
				"Anything that didn't land — a scenario whose folder failed to copy, or a required file (task.md, brief.md, the fixture) absent in the destination. One human-readable line each.",
			items: { type: "string" },
		},
	},
};

const EXERCISE_SCHEMA = {
	type: "object",
	additionalProperties: false,
	required: [
		"completed",
		"summary",
		"docxCommands",
		"otherToolCalls",
		"frictions",
		"outputPath",
	],
	properties: {
		completed: { type: "string", enum: ["yes", "partial", "no"] },
		summary: { type: "string" },
		docxCommands: {
			type: "array",
			description:
				"Every docx-cli invocation you made, in order. This is the docx half of your tool economy.",
			items: {
				type: "object",
				additionalProperties: false,
				required: ["cmd", "outcome"],
				properties: {
					cmd: { type: "string" },
					outcome: { type: "string", enum: ["ok", "error", "confusing"] },
					note: { type: "string" },
				},
			},
		},
		otherToolCalls: {
			type: "integer",
			description:
				"Exact count of NON-docx tool calls you made (file reads, ls/cat, any non-docx shell). docx-cli invocations belong in docxCommands, NOT here. This is the non-docx half of your tool economy.",
		},
		deadEnds: { type: "array", items: { type: "string" } },
		frictions: {
			type: "array",
			items: {
				type: "object",
				additionalProperties: false,
				required: ["issue", "severity"],
				properties: {
					issue: { type: "string" },
					severity: { type: "string", enum: ["blocker", "major", "minor"] },
					suggestion: { type: "string" },
				},
			},
		},
		outputPath: { type: "string" },
	},
};

const RENDER_SCHEMA = {
	type: "object",
	additionalProperties: false,
	required: ["scenarios"],
	properties: {
		scenarios: {
			type: "array",
			items: {
				type: "object",
				additionalProperties: false,
				required: ["key", "rendered"],
				properties: {
					key: { type: "string" },
					rendered: { type: "boolean" },
					pages: { type: "array", items: { type: "string" } },
					baselinePages: { type: "array", items: { type: "string" } },
					error: { type: "string" },
				},
			},
		},
	},
};

const VERDICT_SCHEMA = {
	type: "object",
	additionalProperties: false,
	required: [
		"taskSuccess",
		"rendersCorrectly",
		"survivedReadLoop",
		"merits",
		"defects",
	],
	properties: {
		taskSuccess: { type: "string", enum: ["success", "partial", "fail"] },
		rendersCorrectly: { type: "boolean" },
		formattingPreserved: {
			type: "string",
			enum: ["preserved", "degraded", "broken", "n/a"],
		},
		survivedReadLoop: { type: "boolean" },
		merits: {
			type: "array",
			description:
				"What went RIGHT for this task — tool affordances that worked, things the agent got first-try, parts of the output that are correct/well-formed.",
			items: { type: "string" },
		},
		defects: {
			type: "array",
			items: {
				type: "object",
				additionalProperties: false,
				required: ["what", "severity"],
				properties: {
					what: { type: "string" },
					severity: { type: "string", enum: ["blocker", "major", "minor"] },
					evidence: { type: "string" },
				},
			},
		},
		weakAgentStruggle: { type: "string" },
		notes: { type: "string" },
	},
};

// ---------------------------------------------------------------------------
// Phase 0 — Stage (one agent). Copy ONLY the active scenarios' folders from the
// pristine scenarios dir into the run workspace, one result folder per scenario
// (<runDir>/<key>/). Each scenario folder is self-describing (task.md, brief.md,
// the fixture, assets/), so staging is a plain recursive folder copy — no per-file
// manifest. The skill only makes the empty run dir; the copy lives here because
// workflow scripts can't touch the filesystem, so it runs in an agent. Runs BEFORE
// the exercise token snapshot below, so it doesn't pollute the Haiku measurement.
// ---------------------------------------------------------------------------
phase("Stage");
const stageTargets = active.map((scenario) => ({
	key: scenario.key,
	srcDir: `${scenariosDir}/${scenario.key}`,
	dstDir: `${runDir}/${scenario.key}`,
	requireDoc: scenario.kind === "edit" ? scenario.doc : null,
}));
const stageResult = await agent(stagePrompt(stageTargets), {
	label: "stage:inputs",
	phase: "Stage",
	agentType: "general-purpose",
	schema: STAGE_SCHEMA,
});
const stageMissing = (stageResult && stageResult.missing) || [];
if (stageMissing.length) {
	throw new Error(
		`Staging failed — these inputs never landed in the run dir:\n  ${stageMissing.join("\n  ")}\nCheck that the scenario folders exist under ${scenariosDir}.`,
	);
}
log(
	`Staged ${active.length} scenario folder(s): ${active.map((scenario) => scenario.key).join(", ")}`,
);

// ---------------------------------------------------------------------------
// Phases 1–3 — Exercise → Render → Judge, PIPELINED per scenario. Each scenario
// flows on its own: its Haiku exercise runs in parallel with the others, and the
// MOMENT that exercise finishes its render is enqueued — so the serial render
// queue starts draining as soon as the FIRST exercise completes instead of
// waiting for the slowest. The moment a render finishes, its judge runs. Renders
// are funneled through a 1-slot gate (serializeRender) because Word-mac drives a
// single app instance: two concurrent renders silently export the WRONG document
// (verified empirically). Judges only READ the produced PNGs, so they fan out.
//
// Trade-off: because render/judge now overlap the exercises, we no longer isolate
// the Haiku token cost with a budget.spent() window (other models run in the same
// span). Accurate per-agent tokens + wall-clock time come from the skill's
// transcript pass (scripts/haiku-metrics.ts); the self-reported docx/other tool
// split below is independent of this and unaffected.
// ---------------------------------------------------------------------------
phase("Exercise");

// 1-slot gate that serializes every Word render across the whole pipeline. Reset
// per run; serializeRender (hoisted, below) chains each render behind the previous.
let renderGate = Promise.resolve();

const pipelines = active.map((scenario) => {
	const exerciseP = agent(exercisePrompt(scenario), {
		label: `exercise:${scenario.key}`,
		phase: "Exercise",
		model: "haiku",
		agentType: "general-purpose",
		schema: EXERCISE_SCHEMA,
	})
		.then((result) => (result ? { ...result, key: scenario.key } : null))
		.catch(() => null);

	// Render fires as soon as THIS exercise resolves, queued behind any in-flight
	// render. Skip if the exercise produced nothing.
	const renderP = exerciseP.then((exercise) =>
		exercise ? serializeRender(() => renderOne(scenario)) : null,
	);

	// Judge fires as soon as this scenario's render resolves (it can tolerate a
	// null/failed render — judgePrompt handles that).
	const judgeP = Promise.all([exerciseP, renderP]).then(([exercise, render]) =>
		exercise
			? agent(judgePrompt(scenario, exercise, render), {
					label: `judge:${scenario.key}`,
					phase: "Judge",
					agentType: "general-purpose",
					schema: VERDICT_SCHEMA,
				})
					.then((verdict) => ({ ...verdict, key: scenario.key }))
					.catch(() => null)
			: null,
	);

	return { exerciseP, renderP, judgeP };
});

const exercises = (
	await Promise.all(pipelines.map((pipeline) => pipeline.exerciseP))
).filter(Boolean);
log(
	`Exercise done: ${exercises.map((e) => `${e.key}=${e.completed}`).join(", ")}`,
);

// Tool economy — Haiku agents only, split docx-cli vs everything else. The headline
// benchmark metric: how much of a weak agent's effort docx-cli absorbs versus how
// much it spends working around the tool. (docx/other counts are self-reported per
// agent; the skill's haiku-metrics.ts pass produces the accurate transcript-measured
// counts, tokens, and per-agent time for the report.)
const benchmark = buildBenchmark(exercises);
log(
	`Tool economy (Haiku): ${benchmark.totals.docxCalls} docx + ${benchmark.totals.otherCalls} other = ${benchmark.totals.totalCalls} calls; docx share ${Math.round(benchmark.totals.docxShare * 100)}%`,
);

const renders = (
	await Promise.all(pipelines.map((pipeline) => pipeline.renderP))
).filter(Boolean);
log(
	`Render done: ${renders.map((r) => `${r.key}=${r.rendered ? (r.pages || []).length + "p" : "FAIL"}`).join(", ")}`,
);

const verdicts = (
	await Promise.all(pipelines.map((pipeline) => pipeline.judgeP))
).filter(Boolean);
log(
	`Judge done: ${verdicts.map((v) => `${v.key}=${v.taskSuccess}`).join(", ")}`,
);

// ---------------------------------------------------------------------------
// Phase 4 — Synthesize (one strong agent). Prioritized improvement report.
// ---------------------------------------------------------------------------
phase("Synthesize");
const report = await agent(synthPrompt(active, exercises, verdicts, benchmark), {
	label: "synthesize",
	phase: "Synthesize",
	agentType: "general-purpose",
});

return { report, runDir, binary, exercises, verdicts, benchmark };

// ===========================================================================
// Prompt builders (hoisted function declarations)
// ===========================================================================

function exercisePrompt(scenario) {
	const dir = `${runDir}/${scenario.key}`;
	const workLine =
		scenario.kind === "edit"
			? `Your working document (already a private copy — edit it IN PLACE, do NOT use -o/--output):\n  ${dir}/${scenario.doc}`
			: `You are authoring from scratch. Create your output at EXACTLY this path:\n  ${dir}/${scenario.doc}`;

	return `You are stress-testing **docx-cli**, a command-line tool that lets agents read, edit, and comment on Microsoft Word (.docx) files. You are playing the role of a CAPABLE-BUT-FRESH agent (think Haiku): you have NOT used this tool before. Discover everything you need from the tool's own help — do not assume flags.

The CLI executable is at this absolute path (invoke it directly):
  ${binary}

Start by orienting yourself:
  ${binary} --help
  ${binary} info locators
  ${binary} <command> --help     (for any command you intend to use)

${workLine}

Everything you need is in YOUR scenario folder:
  ${dir}
Read these with the Read tool before you start:
  ${dir}/task.md     — your task and the resolution criteria it will be judged against
  ${dir}/brief.md    — the detailed brief: the data to enter and step-by-step instructions
  ${dir}/assets/     — any additional input files (data, images). \`ls\` it; if it holds files, Read them. It may be empty.

## Your task — ${scenario.bucket}  (scenario: ${scenario.key})
Read ${dir}/task.md and ${dir}/brief.md, then carry the task out on the working document above.

## Rules
- STAY IN YOUR SCENARIO FOLDER. The only document you touch is the working file above; the only other files you read live under ${dir} (task.md, brief.md, assets/). Do NOT search the wider filesystem (no roaming \`find\`, no \`ls\`/\`cat\` of other directories), and do NOT copy files in from elsewhere. The run workspace contains OTHER scenarios' folders with look-alike fixtures that are NOT yours — touching them corrupts the test and wastes calls. If something seems missing, re-read your working file; don't go hunting.
- Use ONLY the docx-cli executable above for document operations. Do NOT hand-edit the XML, unzip the .docx, or reach for any other docx library. The whole point is to test THIS tool.
- You MAY use the Read tool on your task.md / brief.md / assets, and run \`${binary} read <file>\` to inspect your progress.
- Locators (p0, t0:r1c2:p0, sN, etc.) shift after structural edits — re-read when needed. Prefer batch operations where the tool offers them.
- If a command fails or confuses you, try at most ~3 reasonable alternatives, then RECORD it as friction and move on. Do not loop forever on one step.
- Make a genuine, complete attempt. Finish the task if you can.

## What to report (this is the actual product of your run)
Return the structured result. Be brutally honest — surfacing rough edges is the entire purpose:
- completed: yes | partial | no
- summary: one short paragraph of what you actually accomplished.
- docxCommands: EVERY docx-cli invocation you ran, in order, each with outcome (ok | error | confusing) and a brief note (especially WHY something errored or confused you). This is measured — be complete and accurate.
- otherToolCalls: the exact integer count of every NON-docx tool call you made (reading the task/brief/asset files, ls, cat, any non-docx shell). Do NOT count docx-cli runs here — those go in docxCommands. We use docxCommands-count vs otherToolCalls to measure how much of your effort went into docx-cli versus working around it, so count carefully.
- deadEnds: wrong turns, retries, things you expected to work but didn't.
- frictions: concrete "what could have been easier?" points, each with severity (blocker | major | minor) and a suggested fix. Include discoverability gaps (couldn't find the right command/flag), confusing output, and anything that made a weak agent likely to fail.
- outputPath: the absolute path to the .docx you produced (it should be ${dir}/${scenario.doc}).`;
}

function renderPrompt(target) {
	const baselineLine = target.baselineDoc
		? `\n  ALSO render the pristine baseline ${target.baselineDoc} into ${target.baselineOutDir}`
		: "";

	return `You are the RENDER step of an evaluation harness. Render the finished .docx to page PNGs using docx-cli's render command, driven by Microsoft **Word**.

The CLI executable:
  ${binary}

Command shape (confirm with \`${binary} render --help\`):
  ${binary} render <FILE> --engine word --out <DIR>

⚠️ CRITICAL: Render STRICTLY ONE AT A TIME (sequentially). The Word engine drives a single Word application instance and concurrent renders corrupt each other. Never background a render or start a second before the first returns.

Create output directories with \`mkdir -p\` as needed. Render this target (key "${target.key}"):
  render ${target.output} into ${target.outDir}${baselineLine}

Capture the produced PNG page paths (the render command prints them). If a render fails, capture the error text and move on — do not retry more than once.

Return the structured result with ONE entry in \`scenarios\` for key "${target.key}": whether it rendered, the list of output page PNG paths, the list of baseline page PNG paths (empty if none), and any error text.`;
}

function judgePrompt(scenario, exercise, render) {
	const dir = `${runDir}/${scenario.key}`;
	const outputPath =
		(exercise && exercise.outputPath) || `${dir}/${scenario.doc}`;
	const reviewPath = `${dir}/review.md`;
	const docxCommands = (exercise && exercise.docxCommands) || [];
	const otherToolCalls = (exercise && exercise.otherToolCalls) || 0;
	const docxErrors = docxCommands.filter(
		(command) => command.outcome === "error" || command.outcome === "confusing",
	).length;
	const exerciseJson = exercise
		? JSON.stringify(
				{
					completed: exercise.completed,
					summary: exercise.summary,
					frictions: exercise.frictions,
					deadEnds: exercise.deadEnds,
					outputPath: exercise.outputPath,
					// Self-reported tool economy — the docx-cli call log (each with its
					// outcome) plus the non-docx call count. The judge uses this to GROUND
					// its agent-struggle assessment, not to penalize a correct output.
					toolEconomy: {
						docxCalls: docxCommands.length,
						docxErrorsOrConfusing: docxErrors,
						otherToolCalls,
						docxCommands,
					},
				},
				null,
				2,
			)
		: "(the exercise agent returned nothing)";
	const pages = (render && render.pages) || [];
	const baselinePages = (render && render.baselinePages) || [];
	const renderLine = render
		? `Rendered: ${render.rendered}. Output page PNGs:\n${pages.map((p) => `  ${p}`).join("\n") || "  (none)"}${baselinePages.length ? `\nBaseline page PNGs:\n${baselinePages.map((p) => `  ${p}`).join("\n")}` : ""}${render.error ? `\nRender error: ${render.error}` : ""}`
		: "(no render result for this scenario)";

	return `You are a STRICT evaluator judging whether docx-cli let a weak (Haiku) agent complete a real task, and whether the result is correct and well-formed. Be skeptical: a self-reported "completed: yes" means nothing until you verify it.

## Scenario: ${scenario.key} — ${scenario.bucket}
The task given to the weak agent AND the ground-truth resolution criteria are both in:
  ${dir}/task.md
READ that file first — it defines what the agent was asked to do and what "correct" means. (The detailed brief the agent followed is ${dir}/brief.md if you need it.)

## The weak agent's self-report
${exerciseJson}

## Renders produced (Word)
${renderLine}

## How to judge
1. READ ${dir}/task.md for the task + resolution criteria.
2. READ the output page PNG(s) with the Read tool and look at them critically — does the document actually accomplish the task and look right (layout, no leftover placeholders/highlights, tables intact, figure present, columns present, etc.)?
3. Run \`${binary} read ${outputPath}\` to confirm the changes SURVIVE THE WRITE→READ LOOP — this is docx-cli's core invariant; an edit that isn't retrievable on the next read is a failure. Use \`--ast\` if you need structure (e.g. section columns, tracked changes), and \`${binary} track-changes list\` / \`${binary} comments list\` where the scenario calls for them.
4. ${scenario.baseline ? "Compare the BASELINE renders against the OUTPUT renders: confirm ONLY the intended cells changed and all other formatting/headers/footers/structure is preserved." : "Cross-check the render against the criteria."}
5. Weigh the agent's TOOL ECONOMY (in the self-report's \`toolEconomy\`: the docx-cli call log with each call's outcome, plus the non-docx call count). Use it to GROUND your agent-struggle read — many calls or several \`error\`/\`confusing\` outcomes on a simple task is a UX signal; cite the specific commands that errored. IMPORTANT: this informs the agent-struggle / UX dimension ONLY. Do NOT downgrade \`taskSuccess\`, \`rendersCorrectly\`, or \`formattingPreserved\` because the agent thrashed — those are judged purely from the render + the write→read loop. A correct output reached via a painful path is still a task SUCCESS with a UX demerit.
6. Separate two questions: did the AGENT struggle (a UX problem) vs. is the TOOL broken (a bug)? Both matter.

The CLI executable for your verification commands: ${binary}

## Write your review to disk
After you've judged, WRITE a human-readable Markdown review to EXACTLY this path (use the Write tool):
  ${reviewPath}
The review must include: the scenario key + bucket, your verdict (task success, renders correctly, formatting preserved, survived read loop), a **Merits** section (what went right), a **Demerits** section (each defect with its severity and the concrete evidence you saw in the render or read output), and a **Tool economy** section (the self-reported docx-cli call count, how many errored/confused, the non-docx call count, and a one-line read on whether the path to the result was smooth or a slog — naming the specific commands that tripped the agent up). This file is the saved judge's review for this task — make it complete and self-contained. (A precise, transcript-measured per-task metrics file lands next to it after the run; your Tool economy section is the qualitative read on the self-reported numbers.)

Then return the structured verdict. Record BOTH sides for this task:
- merits: what went right (what the tool made easy, what the agent got correct, parts of the output that are well-formed). Always list at least one if anything worked.
- defects: the demerits — concrete, evidence-backed failures (cite what you saw in the render or read output), each with a severity.`;
}

function synthPrompt(scenarios, exercises, verdicts, benchmark) {
	const payload = JSON.stringify(
		{
			scenarios: scenarios.map((scenario) => ({
				key: scenario.key,
				bucket: scenario.bucket,
				kind: scenario.kind,
			})),
			exercises,
			verdicts,
			benchmark,
		},
		null,
		2,
	);

	return `You are writing the final report of an adversarial usability review of **docx-cli**. Weak (Haiku) agents attempted ${scenarios.length} real document tasks; a stricter judge then graded each result against ground truth using Word renders and the write→read loop. Your audience is the engineer who maintains docx-cli. The central question: **can weak agents actually use this tool to get real work done, and what should we fix first?**

Here is all the data (every weak agent's self-report + every judge verdict):
\`\`\`json
${payload}
\`\`\`

Write a thorough, prioritized Markdown report with these sections:

1. **Executive summary** — can weak agents use docx-cli today? Overall pass rate, the headline strengths, and the 2–3 biggest problems.
2. **Scoreboard** — a Markdown table: scenario | bucket | task success (success/partial/fail) | renders correctly | formatting preserved | docx calls | other tool calls | top merit | top demerit. The docx/other call counts come from \`benchmark.perScenario\`.
2b. **Per-task merits & demerits** — for EVERY scenario, a short block listing its merits (what worked) and its demerits (defects/failures) from the judge verdicts. The user explicitly wants both sides for each task.
2c. **Tool economy** — a short subsection on the Haiku tool split from \`benchmark\`: total docx-cli calls vs other tool calls and the docx share, plus per-scenario outliers (which tasks needed the most docx calls or the most non-docx workaround calls) and what that says about ergonomics. A high non-docx share, or many docx calls for a simple task, is a friction signal. Note: the harness appends a precise, transcript-measured metrics table (per-agent tokens + wall-clock time + tool split) below your report — you may reference it, but you don't need to reproduce the exact numbers; the self-reported \`benchmark\` counts you have are approximate.
3. **Cross-cutting themes** — group findings into: Discoverability, CLI ergonomics / surface, Correctness & bugs, Formatting fidelity / preservation, Missing capabilities. Rank themes by impact. For each, give the EVIDENCE (which scenarios, specific commands, judge defects, verbatim friction quotes) and a concrete, actionable recommendation.
4. **Prioritized fixes** — a numbered top 5–8 list, highest leverage first, each tied to the evidence above and phrased as something the maintainer can act on (ideally pointing at the command/flag/output to change).
5. **What worked well** — what weak agents found easy; don't only criticize.

Cite scenario keys throughout and quote agent friction verbatim where it's illuminating. Return the COMPLETE report as your final message — it will be saved to disk and shown to the maintainer.`;
}

// Build the stage agent's prompt: for each active scenario, recursively copy its
// pristine folder into its own run-dir subfolder, then verify the required inputs
// landed. The scenario folder is the unit of staging — task.md, brief.md, the
// fixture, and assets/ all travel together.
function stagePrompt(targets) {
	const lines = targets
		.map((target) => {
			const docLine = target.requireDoc
				? `\n      then verify these exist in the destination: ${target.dstDir}/task.md, ${target.dstDir}/brief.md, ${target.dstDir}/${target.requireDoc}`
				: `\n      then verify these exist in the destination: ${target.dstDir}/task.md, ${target.dstDir}/brief.md   (authoring scenario — no fixture .docx)`;
			return `  - "${target.key}": \`mkdir -p ${target.dstDir}\` then copy the FOLDER CONTENTS: \`cp -R ${target.srcDir}/. ${target.dstDir}/\`${docLine}`;
		})
		.join("\n");

	return `You are the STAGE step of an evaluation harness. Seed an isolated run workspace by copying EXACTLY the listed scenario folders to their destinations and nothing else. Each scenario is a self-contained folder (task.md, brief.md, the fixture .docx, assets/); copy the whole folder so its inputs travel together.

Scenario folders to stage:
${lines}

For each scenario: \`mkdir -p\` the destination, then \`cp -R <srcDir>/. <dstDir>/\` to copy the folder contents (the trailing \`/.\` copies the CONTENTS, including the assets/ subfolder). Sources are pristine — never edit them. Do not copy any folders beyond this list. Do not retry a failed copy more than once.

Return the structured result:
- staged: the destination folders that now exist with their required files present (verify each file with a test/ls before listing the folder).
- missing: one human-readable line for anything that didn't land — a scenario whose copy failed, or a required file (task.md / brief.md / the fixture) absent in the destination, formatted like "key: <what is missing>". Empty array if everything copied.`;
}

// Aggregate the Haiku tool economy from the exercise self-reports: docx-cli calls
// vs every other tool call, per scenario and in total. docxCalls is the length of
// the reported docxCommands log; otherCalls is the agent's otherToolCalls count.
function buildBenchmark(exerciseResults) {
	const perScenario = exerciseResults.map((exercise) => {
		const docxCalls = (exercise.docxCommands || []).length;
		const otherCalls = exercise.otherToolCalls || 0;
		const totalCalls = docxCalls + otherCalls;
		return {
			key: exercise.key,
			completed: exercise.completed,
			docxCalls,
			otherCalls,
			totalCalls,
			docxShare: totalCalls ? Number((docxCalls / totalCalls).toFixed(3)) : 0,
		};
	});
	const docxCalls = perScenario.reduce((acc, row) => acc + row.docxCalls, 0);
	const otherCalls = perScenario.reduce((acc, row) => acc + row.otherCalls, 0);
	const totalCalls = docxCalls + otherCalls;
	return {
		note: "Haiku exercise agents only. Tool counts are self-reported; the skill's haiku-metrics.ts pass over the transcripts has the accurate per-agent counts, tokens, and time (the pipeline overlaps phases, so tokens can't be isolated in-workflow).",
		perScenario,
		totals: {
			scenarios: perScenario.length,
			docxCalls,
			otherCalls,
			totalCalls,
			docxShare: totalCalls ? Number((docxCalls / totalCalls).toFixed(3)) : 0,
		},
	};
}

// Normalize the `only` scenario filter into an array of trimmed keys (or undefined
// for "run everything"). Accepts an array, a single key string ("mnda"), or a
// comma/space-separated string ("mnda, loi") so a caller can run ONE task without
// array ceremony. Unknown keys are caught downstream where `active` ends up empty.
function normalizeOnly(value) {
	if (!value) return undefined;
	const keys = Array.isArray(value) ? value : String(value).split(/[\s,]+/);
	const cleaned = keys.map((key) => String(key).trim()).filter(Boolean);
	return cleaned.length ? cleaned : undefined;
}

// 1-slot async gate serializing every Word render across the pipeline: Word-mac is
// a single app instance, so renders must run strictly one at a time even though
// exercises and judges fan out. Each call queues behind the previous render; the
// gate advances regardless of success/failure so one bad render can't wedge the
// rest. (`renderGate` is the per-run `let` declared next to the pipeline.)
function serializeRender(thunk) {
	const run = renderGate.then(thunk, thunk);
	renderGate = run.then(
		() => undefined,
		() => undefined,
	);
	return run;
}

// Render ONE scenario's output (plus its baseline, if any) into that scenario's
// result folder (<runDir>/<key>/renders/{output,baseline}/) and return the render
// record. A failed render degrades to a rendered:false record the judge can still
// handle rather than rejecting.
function renderOne(scenario) {
	const dir = `${runDir}/${scenario.key}`;
	const target = {
		key: scenario.key,
		output: `${dir}/${scenario.doc}`,
		outDir: `${dir}/renders/output/`,
		baselineDoc: scenario.baseline
			? `${scenariosDir}/${scenario.key}/${scenario.doc}`
			: null,
		baselineOutDir: scenario.baseline ? `${dir}/renders/baseline/` : null,
	};
	return agent(renderPrompt(target), {
		label: `render:${scenario.key}`,
		phase: "Render",
		agentType: "general-purpose",
		schema: RENDER_SCHEMA,
	})
		.then((result) => {
			const list = (result && result.scenarios) || [];
			return (
				list.find((entry) => entry.key === scenario.key) ||
				list[0] || { key: scenario.key, rendered: false, error: "no render result" }
			);
		})
		.catch((error) => ({
			key: scenario.key,
			rendered: false,
			error: String((error && error.message) || error),
		}));
}
