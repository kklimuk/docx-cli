export const meta = {
	name: "adversarial-review",
	description:
		"Weak-agent (Haiku) adversarial test of docx-cli: 8 scenarios over real fixtures + authoring tasks, rendered with Word, judged, and synthesized into a prioritized ergonomics report.",
	phases: [
		{
			title: "Stage",
			detail:
				"copy only the active scenarios' inputs from fixtures into the run dir",
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
				"a strong agent grades each render + the Haiku transcript, firing as soon as that render is done",
		},
		{
			title: "Synthesize",
			detail: "one strong agent writes the prioritized improvement report",
		},
	],
};

// ---------------------------------------------------------------------------
// args, injected by the skill: { runDir, binary, fixturesDir, only? }
//   runDir       — absolute /tmp dir holding a working COPY of the fixtures (docx + assets/)
//   binary       — absolute path to the freshly-built dist/docx
//   fixturesDir  — absolute path to the PRISTINE bundled fixtures (for render baselines)
//   only         — optional scenario filter: run just these key(s). Accepts an
//                  array (["mnda","loi"]), a single key ("mnda" — the natural way
//                  to run ONE task), or a comma/space-separated string ("mnda,loi").
// ---------------------------------------------------------------------------
// The runtime delivers `args` as a JSON STRING (not the parsed object the tool
// docs imply), so parse it if needed. Accept an already-parsed object too.
const parsedArgs = typeof args === "string" ? JSON.parse(args) : args || {};
const { runDir, binary, fixturesDir } = parsedArgs;
const only = normalizeOnly(parsedArgs.only);
if (!runDir || !binary || !fixturesDir) {
	throw new Error(
		"adversarial-review requires args { runDir, binary, fixturesDir }",
	);
}

const SCENARIOS = [
	{
		key: "mnda",
		bucket: "Form filling + highlight removal",
		kind: "edit",
		work: "mnda.docx",
		fixture: "mnda.docx",
		brief: "assets/mnda-deal.md",
		out: "mnda.docx",
		baseline: false,
		task: "Fill out the MNDA cover page from the deal sheet and REMOVE every yellow highlight (the placeholders were highlighted; once filled, no highlight should remain anywhere). Edit the file in place.",
		criteria:
			"Cover-page placeholders replaced with the real party names, dates, term (2y), confidentiality term (3y), Delaware governing law / New Castle County jurisdiction; signature table populated for both parties; ZERO yellow highlight left in the document. Changes must be visible via `docx read`.",
	},
	{
		key: "loi",
		bucket: "Comments + tracked changes (legal review)",
		kind: "edit",
		work: "loi.docx",
		fixture: "loi.docx",
		brief: "assets/loi-review.md",
		out: "loi.docx",
		baseline: false,
		task: "Act as counsel for the Customer. With tracked changes ON, delete the four unwanted alternative Confidentiality clauses (keep only the plain 'Mutual' one) as tracked deletions, and leave anchored comments on the unfilled blanks, the kept clause, and the one-sided investor-disclosure carve-out.",
		criteria:
			"Tracked changes enabled; the 4 redundant confidentiality variants removed as tracked deletions (redlines still visible, not silently gone); the 'Mutual' variant kept; comments anchored to the relevant text for the blanks, the kept clause, and the investor carve-out. `docx comments list` and `docx track-changes list` should reflect this.",
	},
	{
		key: "invoice",
		bucket: "Table editing + embedded-image preservation",
		kind: "edit",
		work: "invoice.docx",
		fixture: "invoice.docx",
		brief: "assets/invoice-data.md",
		out: "invoice.docx",
		baseline: false,
		task: "Fill the invoice's three tables with the supplied company/customer details, three line items, and totals. Keep the layout, the embedded logo image, and the formatting intact.",
		criteria:
			"All placeholder cells (company, customer, Item 1/2/3, $0.00 amounts, dates, invoice #) replaced with the supplied values; the three tables still well-formed; the embedded logo image still present (the doc has two media images). Totals match the data sheet. Changes visible via `docx read`.",
	},
	{
		key: "psa",
		bucket: "Formatting + unmodeled-structure preservation",
		kind: "edit",
		work: "professional-services-agreement.docx",
		fixture: "professional-services-agreement.docx",
		brief: "assets/psa-preservation.md",
		out: "professional-services-agreement.docx",
		baseline: true,
		task: "With tracked changes ON, make ONLY two surgical edits — fill the 'Payment Period' cell ('30 days from Customer's receipt of invoice') and the 'Invoice Period' cell ('Monthly') — and leave the entire rest of this large, heavily-formatted contract (drafting notes, checkboxes, 10 headers/footers, attached Standard Terms) untouched.",
		criteria:
			"`docx track-changes list` shows EXACTLY the two intended tracked edits and nothing else. Before/after renders are visually identical everywhere except those two cells. All gray drafting notes, checkboxes, every header/footer, and the full Standard Terms part survive unchanged. This is the core in-place-mutation fidelity test.",
	},
	{
		key: "business-letter",
		bucket: "Editing / mail-merge (real-world use case)",
		kind: "edit",
		work: "business-letter.docx",
		fixture: "business-letter.docx",
		brief: "assets/business-letter-brief.md",
		out: "business-letter.docx",
		baseline: false,
		task: "Replace every {{Mustache}} token in the letter with the supplied real value so it reads as a finished, signable letter. Leave no {{ }} tokens behind; preserve the bold/colored runs around amounts and the sender name and the two enclosure bullets.",
		criteria:
			"Every {{token}} replaced with the correct value from the brief; no mustache tokens remain; bold/colored runs and enclosure bullets preserved. Changes visible via `docx read`.",
	},
	{
		key: "resume",
		bucket: "Styling fidelity + drawing preservation",
		kind: "edit",
		work: "resume.docx",
		fixture: "resume.docx",
		brief: "assets/resume-candidate.md",
		out: "resume.docx",
		baseline: false,
		task: "Fill the Harvard résumé template for the supplied candidate, preserving heading styles, the right-aligning tab stops on dates, and the bullet lists. Remove the bracketed [Note: ...] helper text. Leave the [drawing] element alone.",
		criteria:
			"Name, contact, education, experience, leadership filled from the candidate sheet; [Note: ...] helpers gone; section headings keep their Heading style; dates still right-align at the tab; the [drawing] element survives. Render should look like a clean résumé.",
	},
	{  
		key: "papers-report",
		bucket: "Authoring: sections, headings, citations, embedded figure",
		kind: "author",
		brief: "assets/papers/report-brief.md",
		assets: [
			"assets/papers/paper-1.md",
			"assets/papers/paper-2.md",
			"assets/papers/paper-3.md",
			"assets/papers/figure-1.png",
		],
		out: "report.docx",
		baseline: false,
		task: "Author a NEW report.docx literature review from the three papers: title + intro, one styled-heading section per paper with an in-text citation, a Comparison section that EMBEDS figure-1.png with a caption, and a References list.",
		criteria:
			"report.docx exists with real heading styles, three paper summaries each carrying an (Author, year) citation, the figure-1.png chart actually embedded and rendering, a caption, and a references section. Multi-section structure renders cleanly.",
	},
	{
		key: "eliot-journal",
		bucket: "Authoring: Markdown, multi-column sections, verse",
		kind: "author",
		brief: "assets/eliot-journal-brief.md",
		assets: ["assets/eliot-poems.md"],
		out: "journal.docx",
		baseline: false,
		task: "Author a NEW journal.docx T. S. Eliot poetry reader: a single-column title section, a TWO-COLUMN body section holding the poems with their titles as headings and verse line breaks preserved, and a closing single-column colophon section.",
		criteria:
			"journal.docx exists; a two-column section actually holds the poems (verify via `docx read --ast` sectPr columns or the render showing side-by-side columns); poem titles are headings; verse line breaks are preserved (lines NOT collapsed into one paragraph); title and colophon are single-column.",
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
			description: "Destination paths that now exist after copying.",
			items: { type: "string" },
		},
		missing: {
			type: "array",
			description:
				"Destinations whose source was absent or whose copy failed — each as 'src -> dst'.",
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
// Phase 0 — Stage (one agent). Copy ONLY the active scenarios' inputs from the
// pristine fixtures into the run workspace. SCENARIOS is the single source of
// truth for what each scenario needs, so a single-scenario run stages just that
// scenario's files instead of the whole corpus. The skill only makes the empty
// run dir; the copy lives here because workflow scripts can't touch the
// filesystem, so it runs in an agent. Runs BEFORE the exercise token snapshot
// below, so it doesn't pollute the Haiku token measurement.
// ---------------------------------------------------------------------------
phase("Stage");
const stageCopies = buildStageCopies(active, fixturesDir, runDir);
const stageResult = await agent(stagePrompt(stageCopies), {
	label: "stage:inputs",
	phase: "Stage",
	agentType: "general-purpose",
	schema: STAGE_SCHEMA,
});
const stageMissing = (stageResult && stageResult.missing) || [];
if (stageMissing.length) {
	throw new Error(
		`Staging failed — these inputs never landed in the run dir:\n  ${stageMissing.join("\n  ")}\nCheck that the fixtures exist under ${fixturesDir}.`,
	);
}
log(
	`Staged ${stageCopies.length} input file(s) for ${active.length} scenario(s): ${active.map((scenario) => scenario.key).join(", ")}`,
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
// transcript pass (scripts/haiku-metrics.py); the self-reported docx/other tool
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
// agent; the skill's haiku-metrics.py pass produces the accurate transcript-measured
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
	const briefPath = `${runDir}/${scenario.brief}`;
	const extraAssets = (scenario.assets || [])
		.map((relativePath) => `${runDir}/${relativePath}`)
		.filter((path) => path !== briefPath);
	const workLine =
		scenario.kind === "edit"
			? `Your working document (already a private copy — edit it IN PLACE, do NOT use -o/--output):\n  ${runDir}/${scenario.work}`
			: `You are authoring from scratch. Create your output at EXACTLY this path:\n  ${runDir}/${scenario.out}`;
	const assetLines = [briefPath, ...extraAssets]
		.map((path) => `  ${path}`)
		.join("\n");

	return `You are stress-testing **docx-cli**, a command-line tool that lets agents read, edit, and comment on Microsoft Word (.docx) files. You are playing the role of a CAPABLE-BUT-FRESH agent (think Haiku): you have NOT used this tool before. Discover everything you need from the tool's own help — do not assume flags.

The CLI executable is at this absolute path (invoke it directly):
  ${binary}

Start by orienting yourself:
  ${binary} --help
  ${binary} info locators
  ${binary} <command> --help     (for any command you intend to use)

${workLine}

Read these files for your task data and instructions (use the Read tool):
${assetLines}

## Your task — ${scenario.bucket}  (scenario: ${scenario.key})
${scenario.task}

Follow the detailed brief in ${briefPath}.

## Rules
- STAY ON YOUR FILE. The only document you touch is the working file above; the only other files you read are the asset files listed above. Do NOT search the wider filesystem (no roaming \`find\`, no \`ls\`/\`cat\` of other directories), and do NOT copy files in from elsewhere. The workspace contains other scenarios' look-alike fixtures and earlier run outputs that are NOT yours — touching them corrupts the test and wastes calls. If something seems missing, re-read the working file; don't go hunting.
- Use ONLY the docx-cli executable above for document operations. Do NOT hand-edit the XML, unzip the .docx, or reach for any other docx library. The whole point is to test THIS tool.
- You MAY use the Read tool on the asset files and on any images, and run \`${binary} read <file>\` to inspect your progress.
- Locators (p0, t0:r1c2:p0, sN, etc.) shift after structural edits — re-read when needed. Prefer batch operations where the tool offers them.
- If a command fails or confuses you, try at most ~3 reasonable alternatives, then RECORD it as friction and move on. Do not loop forever on one step.
- Make a genuine, complete attempt. Finish the task if you can.

## What to report (this is the actual product of your run)
Return the structured result. Be brutally honest — surfacing rough edges is the entire purpose:
- completed: yes | partial | no
- summary: one short paragraph of what you actually accomplished.
- docxCommands: EVERY docx-cli invocation you ran, in order, each with outcome (ok | error | confusing) and a brief note (especially WHY something errored or confused you). This is measured — be complete and accurate.
- otherToolCalls: the exact integer count of every NON-docx tool call you made (reading the brief/asset files, ls, cat, any non-docx shell). Do NOT count docx-cli runs here — those go in docxCommands. We use docxCommands-count vs otherToolCalls to measure how much of your effort went into docx-cli versus working around it, so count carefully.
- deadEnds: wrong turns, retries, things you expected to work but didn't.
- frictions: concrete "what could have been easier?" points, each with severity (blocker | major | minor) and a suggested fix. Include discoverability gaps (couldn't find the right command/flag), confusing output, and anything that made a weak agent likely to fail.
- outputPath: the absolute path to the .docx you produced.`;
}

function renderPrompt(targets) {
	const lines = targets
		.map((target) => {
			const base = target.baseline
				? `\n    ALSO render the pristine baseline ${target.baseline} into ${runDir}/render/${target.key}/baseline/`
				: "";
			return `  - key "${target.key}": render ${target.output} into ${runDir}/render/${target.key}/output/${base}`;
		})
		.join("\n");

	return `You are the RENDER step of an evaluation harness. Render each finished .docx to page PNGs using docx-cli's render command, driven by Microsoft **Word**.

The CLI executable:
  ${binary}

Command shape (confirm with \`${binary} render --help\`):
  ${binary} render <FILE> --engine word --out <DIR>

⚠️ CRITICAL: Render STRICTLY ONE AT A TIME (sequentially). The Word engine drives a single Word application instance and concurrent renders corrupt each other. Never background a render or start a second before the first returns.

Create output directories with \`mkdir -p\` as needed. Render these targets:
${lines}

For each target, capture the produced PNG page paths (the render command prints them). If a render fails, capture the error text and move on to the next target — do not retry more than once.

Return the structured result: for every key, whether it rendered, the list of output page PNG paths, the list of baseline page PNG paths (empty if none), and any error text.`;
}

function judgePrompt(scenario, exercise, render) {
	const exerciseJson = exercise
		? JSON.stringify(
				{
					completed: exercise.completed,
					summary: exercise.summary,
					frictions: exercise.frictions,
					deadEnds: exercise.deadEnds,
					outputPath: exercise.outputPath,
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
Task given to the weak agent:
${scenario.task}

## Ground-truth success criteria (what "correct" means)
${scenario.criteria}

## The weak agent's self-report
${exerciseJson}

## Renders produced (Word)
${renderLine}

## How to judge
1. READ the output page PNG(s) with the Read tool and look at them critically — does the document actually accomplish the task and look right (layout, no leftover placeholders/highlights, tables intact, figure present, columns present, etc.)?
2. Run \`${binary} read ${exercise && exercise.outputPath ? exercise.outputPath : `${runDir}/${scenario.out}`}\` to confirm the changes SURVIVE THE WRITE→READ LOOP — this is docx-cli's core invariant; an edit that isn't retrievable on the next read is a failure. Use \`--ast\` if you need structure (e.g. section columns, tracked changes), and \`${binary} track-changes list\` / \`${binary} comments list\` where the scenario calls for them.
3. ${scenario.baseline ? "Compare the BASELINE renders against the OUTPUT renders: confirm ONLY the intended cells changed and all other formatting/headers/footers/structure is preserved." : "Cross-check the render against the criteria."}
4. Separate two questions: did the AGENT struggle (a UX problem) vs. is the TOOL broken (a bug)? Both matter.

The CLI executable for your verification commands: ${binary}

Return the structured verdict. Record BOTH sides for this task:
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

// Build the scoped copy list for staging: for each active scenario, the input
// files it needs (its source docx for edit scenarios, its brief, and any extra
// assets), as { src in fixtures } -> { dst in run dir } preserving relative paths.
// Deduped by destination so a brief shared across scenarios is copied once.
function buildStageCopies(scenarios, fixturesRoot, runRoot) {
	const seen = new Set();
	const copies = [];
	const add = (relativeSrc, relativeDst) => {
		if (!relativeSrc || seen.has(relativeDst)) return;
		seen.add(relativeDst);
		copies.push({
			src: `${fixturesRoot}/${relativeSrc}`,
			dst: `${runRoot}/${relativeDst}`,
		});
	};
	for (const scenario of scenarios) {
		// edit scenarios get a private working copy of their source docx; authoring
		// scenarios create their output fresh, so they need no docx staged.
		if (scenario.kind === "edit") add(scenario.fixture, scenario.work);
		add(scenario.brief, scenario.brief);
		for (const asset of scenario.assets || []) add(asset, asset);
	}
	return copies;
}

function stagePrompt(copies) {
	const lines = copies.map((copy) => `  ${copy.src}  ->  ${copy.dst}`).join("\n");
	return `You are the STAGE step of an evaluation harness. Copy EXACTLY the listed source files to their destination paths and nothing else — this seeds an isolated run workspace with only the inputs the selected scenarios need.

Copies (source  ->  destination):
${lines}

For each pair: \`mkdir -p\` the destination's PARENT directory, then \`cp\` the source file to the destination. Sources are pristine fixtures — never edit them. Do not copy any files beyond this list. Do not retry a failed copy more than once.

Return the structured result:
- staged: the destination paths that now exist (verify each with a test/ls before listing it).
- missing: any pair, formatted "src -> dst", whose source file was absent or whose copy failed. Empty array if everything copied.`;
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
		note: "Haiku exercise agents only. Tool counts are self-reported; the skill's haiku-metrics.py pass over the transcripts has the accurate per-agent counts, tokens, and time (the pipeline overlaps phases, so tokens can't be isolated in-workflow).",
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

// Render ONE scenario's output (plus its baseline, if any) and return that
// scenario's render record. Reuses the multi-target render prompt/schema with a
// single target and pulls the one row back out, so a failed render degrades to a
// rendered:false record the judge can still handle rather than rejecting.
function renderOne(scenario) {
	const target = {
		key: scenario.key,
		output: `${runDir}/${scenario.out}`,
		baseline: scenario.baseline ? `${fixturesDir}/${scenario.fixture}` : null,
	};
	return agent(renderPrompt([target]), {
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
