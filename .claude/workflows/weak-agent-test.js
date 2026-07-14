export const meta = {
	name: "weak-agent-test",
	description:
		"Weak-agent adversarial test of docx-cli: 6 scenarios over real fixtures + an authoring task, exercised by a swappable weak model (haiku default, sonnet to probe, or a local harness's pre-produced results), rendered with Word, judged and synthesized by fable.",
	phases: [
		{
			title: "Stage",
			detail:
				"stage each active scenario via scripts/stage-scenario.ts (or load the local harness's pre-produced exercise.json results)",
		},
		{
			title: "Exercise",
			detail:
				"one weak agent per scenario performs its task with docx-cli (model: args.model — haiku by default)",
			model: "haiku",
		},
		{
			title: "Render",
			detail:
				"the moment each task finishes, render its output to page PNGs via Word AND save its markdown read view (read.md) — serialized within the run; the CLI's cross-process Word lock makes concurrent runs safe",
		},
		{
			title: "Judge",
			detail:
				"opus grades each render + the exercise account against the rubric and the write→read loop, writing review.md into the scenario's result folder",
			model: "opus",
		},
		{
			title: "Synthesize",
			detail: "opus writes the prioritized improvement report",
			model: "opus",
		},
	],
};

// ---------------------------------------------------------------------------
// args, injected by the skill: { runDir, binary, scenariosDir, only?, model?,
//                                exerciseBackend?, modelLabel?, exercises?,
//                                arm?, competitorDir? }
//   runDir        — absolute run dir; the workflow stages one subfolder PER
//                   active scenario into it (<runDir>/<key>/…) and that subfolder
//                   doubles as the scenario's result folder.
//   binary        — absolute path to the freshly-built dist/docx
//   scenariosDir  — absolute path to the PRISTINE bundled scenarios dir. Each
//                   scenario is a folder named after its key, containing task.md
//                   (the full request, agent-facing), criteria.md (grading rubric,
//                   JUDGE-ONLY — withheld from the agent's run workspace by the stage
//                   step, read by the judge from here), the fixture .docx (edit
//                   scenarios only), and assets/ (extra inputs). The skill's scripts/
//                   dir (stage-scenario.ts) is resolved as its sibling.
//   only          — optional scenario filter: run just these key(s). Accepts an
//                   array (["mnda","loi"]), a single key ("mnda" — the natural way
//                   to run ONE task), or a comma/space-separated string ("mnda,loi").
// ---------------------------------------------------------------------------
// The runtime delivers `args` as a JSON STRING (not the parsed object the tool
// docs imply), so parse it if needed. Accept an already-parsed object too.
const parsedArgs = typeof args === "string" ? JSON.parse(args) : args || {};
const { runDir, binary, scenariosDir } = parsedArgs;
const only = normalizeOnly(parsedArgs.only);
// The exercise (subject-under-test) agent model — SWAPPABLE. Defaults to "haiku" —
// the whole harness is framed around weak agents — but pass args.model: "sonnet" to
// probe a stronger model: does it ACT on the read-time hints/cures a weaker agent
// ignores? Only the exercise agents change; stage/render/judge/synth are unaffected.
const exerciseModel = parsedArgs.model || "haiku";
// The judge and the synthesizer are pinned to the STRONGEST model — grading and
// prioritization are where quality matters most, and they must not drift with the
// session model or the exercise model.
const JUDGE_MODEL = "opus";
const SYNTH_MODEL = "opus";
// Exercise backend: "claude" (default — spawn a Haiku/Sonnet subagent per scenario)
// or "local" (the LOCAL agent harness ran the exercises OUT OF BAND via
// scripts/run-local-corpus.ts, leaving <runDir>/<key>/exercise.json; this workflow
// loads those and only renders/judges/synthesizes them). The local path SKIPS
// Stage + Exercise — the corpus runner already staged each folder (via the SAME
// scripts/stage-scenario.ts the Stage phase uses) and left the worked doc + its
// ledger-measured result. Render/Judge/Synthesize are backend-agnostic and consume
// the same shape, so both backends are graded identically — that is what makes the
// local numbers comparable to Haiku's.
const exerciseBackend = parsedArgs.exerciseBackend || "claude";
const usingLocal = exerciseBackend === "local";
const preExercises = Array.isArray(parsedArgs.exercises) ? parsedArgs.exercises : [];
const preByKey = new Map(preExercises.map((entry) => [entry.key, entry]));
// Human label for the subject-under-test model, spliced into judge/synth prompts.
const modelLabel = usingLocal
	? parsedArgs.modelLabel || "the local harness"
	: exerciseModel;
// Both backends get a MEASURED run-metrics rollup after the run
// (scripts/exercise-metrics.ts): the Claude arms from the agent transcripts, the
// local arm from each scenario's exercise.json `_local` block. Same tables (tokens,
// wall-clock, tool split, correctness), so the judge/synth promise the same thing —
// don't estimate any of it. (`measuredSource` just names where it comes from.)
const measuredSource = usingLocal ? "the harness ledger" : "the run transcripts";
const judgeMetricsNote = `(A measured per-task metrics file — tool calls, tokens, time — is written from ${measuredSource} after the run; do not guess at those numbers.)`;
const synthMetricsNote = `Do NOT invent tool-call counts, token numbers, or timings — the harness appends a MEASURED "Run metrics" section (from ${measuredSource}) below your report after the run; where cost/effort matters to a point you're making, refer the reader to that section instead of estimating.`;
if (!runDir || !binary || !scenariosDir) {
	throw new Error(
		"weak-agent-test requires args { runDir, binary, scenariosDir }",
	);
}
// The skill's scripts live next to its scenarios — stage-scenario.ts is the one
// staging path (the Stage agent runs it; run-local-corpus.ts imports it).
const scriptsDir = `${String(scenariosDir).replace(/\/scenarios\/?$/, "")}/scripts`;

// A/B arm: "docx-cli" (DEFAULT — the existing single-tool harness, unchanged) or
// "anthropic-docx-skill" (the competitor: Anthropic's bundled python/raw-OOXML docx
// skill). Only the EXERCISE agents' tool instructions differ between arms; staging,
// rendering, judging, and metrics are tool-agnostic, so the two arms are graded the
// same way and the outputs compare apples-to-apples. The competitor arm needs
// competitorDir — the staged Anthropic docx skill folder (SKILL.md + scripts/),
// provisioned with its deps by scripts/stage-competitor.ts.
const arm = parsedArgs.arm || "docx-cli";
const competitorDir = parsedArgs.competitorDir || null;
// Everything that differs between arms lives HERE — toolName plus the two prompt
// fragments the judge/synth builders splice in. A new arm is one object literal; the
// builders just read ARMS[arm].* (no scattered `arm === …` ternaries at the use sites).
const ARMS = {
	"docx-cli": { toolName: "docx-cli", judgeNote: "", synthClause: "" },
	"anthropic-docx-skill": {
		toolName: "the Anthropic docx skill",
		// The competitor agent used the Anthropic docx skill, not docx-cli — so the judge
		// grades the OUTPUT only, treating docx-cli's `read` as a neutral inspector of the
		// produced .docx (it is NOT the tool under test here).
		judgeNote: `\n\n**Arm note — competitor run:** the agent did this task with **the Anthropic docx skill**, NOT docx-cli. Judge it from the **Word RENDERS + the criteria.md** — those are the primary, authoritative evidence here. \`taskSuccess\`, \`rendersCorrectly\`, and \`formattingPreserved\` MUST be decided from the rendered pages read against the rubric. The \`${binary} read\` step below is OPTIONAL for this arm and only a convenience inspector — if docx-cli cannot read a file that Word renders correctly, that is NOT a competitor failure (it reflects docx-cli's reader, which is not the tool under test), so do NOT lower any verdict (including \`survivedReadLoop\`) on that basis. Set \`survivedReadLoop\` from whether the intended content is actually present in the render.`,
		synthClause:
			" (the competitor arm — this report characterizes how the Anthropic docx skill fared on the same tasks docx-cli is graded on)",
	},
};
if (!ARMS[arm]) {
	throw new Error(
		`Unknown arm "${arm}". Known arms: ${Object.keys(ARMS).join(", ")}.`,
	);
}
if (arm === "anthropic-docx-skill" && !competitorDir) {
	throw new Error(
		"arm 'anthropic-docx-skill' requires args.competitorDir — the staged Anthropic docx skill folder (run scripts/stage-competitor.ts first).",
	);
}
const toolName = ARMS[arm].toolName;

// Orchestration manifest, keyed by scenario folder name. The CONTENT of each
// scenario — its request, grading rubric, fixture, and extra assets —
// lives in `<scenariosDir>/<key>/` (task.md, criteria.md, the .docx, assets/), NOT
// here. This manifest holds only what the workflow needs to ROUTE:
//   key       — folder name; also the result-folder name under runDir.
//   bucket    — human label for the scenario's category (prompt headers, scoreboard).
//   kind      — "edit" (work a staged copy of `doc` in place) | "author" (create `doc` fresh).
//   doc       — the .docx filename inside the scenario folder. For edit scenarios
//               it is the pristine fixture (staged + edited in place AND rendered as
//               the "before" baseline). For author scenarios it does NOT exist in
//               the scenario folder — the agent creates it in the run dir.
// Whether a scenario gets a BASELINE (before/after) render is DERIVED, not a field:
// an edit scenario always has a pristine source (the fixture at
// <scenariosDir>/<key>/<doc>), so it always gets a baseline; an author scenario
// (eliot-journal) has no source, so it never does. See hasBaseline().
// Mirror any change here into the MANIFEST in scripts/run-local-corpus.ts.
const SCENARIOS = [
	{
		key: "mnda",
		bucket: "Form filling + highlight removal + font fidelity",
		kind: "edit",
		doc: "mnda.docx",
	},
	{
		key: "invoice",
		bucket: "Table editing + restructure + image replace",
		kind: "edit",
		doc: "invoice.docx",
	},
	{
		key: "resume",
		bucket: "Styling fidelity + drawing preservation",
		kind: "edit",
		doc: "resume.docx",
	},
	{
		key: "contract-markup",
		bucket: "Legal review: redlining + commenting",
		kind: "edit",
		doc: "contract.docx",
	},
	{
		key: "contract-finalize",
		bucket: "Legal review: accept/reject + resolve comments",
		kind: "edit",
		doc: "contract-redlined.docx",
	},
	{
		key: "eliot-journal",
		bucket: "Authoring: columns, verse, footnotes, links, figure",
		kind: "author",
		doc: "journal.docx",
	},
];

// A scenario gets a baseline (before/after) render iff it has a pristine source doc
// to render — i.e. every EDIT scenario (its fixture is the "before"); author
// scenarios have nothing to compare against.
function hasBaseline(scenario) {
	return scenario.kind === "edit";
}

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
log(`Arm: ${arm} (tool under test: ${toolName}); exercise: ${modelLabel}`);
if (arm === "anthropic-docx-skill") {
	log(`Competitor skill dir: ${competitorDir}`);
}
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
			description:
				"Scenario keys whose stage-scenario.ts run reported staged:true.",
			items: { type: "string" },
		},
		missing: {
			type: "array",
			description:
				"Every `missing` line reported by a stage-scenario.ts run, plus any scenario whose staging command itself failed. One human-readable line each.",
			items: { type: "string" },
		},
	},
};

// Local backend only: the Stage-slot LOAD agent reads each scenario's pre-produced
// exercise.json off disk and returns them here (verbatim). additionalProperties:true
// so the ledger-measured extras (docxCommands, the `_local` block) ride along
// harmlessly — the judge is handed only the same fields the Claude arm reports.
const LOAD_SCHEMA = {
	type: "object",
	additionalProperties: false,
	required: ["exercises"],
	properties: {
		exercises: {
			type: "array",
			items: {
				type: "object",
				additionalProperties: true,
				// `status` is the local exercise's process-lifecycle field
				// (completed|failed). `completed` is only kept optional for backward
				// compatibility with exercise.json files produced before the switch —
				// don't require it.
				required: ["key", "outputPath"],
				properties: {
					key: { type: "string" },
					status: { type: "string" },
					completed: { type: "string" },
					summary: { type: "string" },
					deadEnds: { type: "array", items: { type: "string" } },
					frictions: { type: "array", items: { type: "object", additionalProperties: true } },
					outputPath: { type: "string" },
				},
			},
		},
	},
};

// What an exercise agent reports back. Deliberately NO tool-call tallies: agents
// under-count their own calls ~2×, so the harness measures tool economy and tokens
// from the transcripts after the run (scripts/exercise-metrics.ts) — the agent's
// job here is the qualitative account (what happened, what hurt), not bookkeeping.
const EXERCISE_SCHEMA = {
	type: "object",
	additionalProperties: false,
	required: ["completed", "summary", "frictions", "outputPath"],
	properties: {
		completed: { type: "string", enum: ["yes", "partial", "no"] },
		summary: { type: "string" },
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
					markdownPath: {
						type: "string",
						description:
							"Absolute path to the saved markdown read view (read.md) of the OUTPUT doc, or empty if `docx read` failed.",
					},
					baselineMarkdownPath: {
						type: "string",
						description:
							"Absolute path to the saved markdown read view (read.md) of the pristine BASELINE doc, or empty if there's no baseline or the read failed.",
					},
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
// Phase 0 — Stage (one agent). Stage ONLY the active scenarios' folders from the
// pristine scenarios dir into the run workspace, one result folder per scenario
// (<runDir>/<key>/), by running scripts/stage-scenario.ts per scenario — the SAME
// script the local corpus runner uses, so every backend stages identically. The
// script copies the folder, strips the judge-only criteria.md, and verifies; the
// agent just runs it and relays the verdicts (workflow scripts can't touch the
// filesystem, so an agent carries the commands).
// ---------------------------------------------------------------------------
phase("Stage");
if (usingLocal) {
	// The local corpus runner already staged each folder AND ran the exercise out of
	// band, leaving <runDir>/<key>/exercise.json. Re-staging would clobber the worked
	// docs, so instead of a Stage copy we take those results here. TWO paths:
	//   • DETERMINISTIC (preferred): the caller passes args.exercises, read straight
	//     off disk by scripts/collect-exercises.ts — the code-computed `status` and the
	//     rest reach the workflow without a model in the loop.
	//   • FALLBACK: no args.exercises, so a LOAD agent (an LLM) reads the files and
	//     returns them — used only because the workflow's JS has no filesystem access.
	//     The status is still code-computed on disk; this hop just re-reads it.
	if (preByKey.size === 0) {
		const loadTargets = active.map((scenario) => ({
			key: scenario.key,
			path: `${runDir}/${scenario.key}/exercise.json`,
		}));
		const loaded = await agent(loadPrompt(loadTargets), {
			label: "load:local-exercises",
			phase: "Stage",
			agentType: "general-purpose",
			schema: LOAD_SCHEMA,
		});
		for (const entry of (loaded && loaded.exercises) || []) {
			if (entry && entry.key) preByKey.set(entry.key, entry);
		}
	}
	if (preByKey.size === 0) {
		throw new Error(
			`Local backend: no exercise results — pass args.exercises or ensure ${runDir}/<key>/exercise.json exists (run scripts/run-local-corpus.ts first).`,
		);
	}
	log(
		`Local backend: loaded ${preByKey.size} pre-produced result(s): ${[...preByKey.keys()].join(", ")}`,
	);
} else {
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
}

// ---------------------------------------------------------------------------
// Phases 1–3 — Exercise → Render → Judge, PIPELINED per scenario. Each scenario
// flows on its own: its exercise runs in parallel with the others, and the
// MOMENT that exercise finishes, its render is enqueued — so the render queue
// starts draining as soon as the FIRST exercise completes instead of waiting for
// the slowest. The moment a render finishes, its judge runs. This holds for BOTH
// arms: a Claude exercise resolves when the agent finishes; a local exercise
// resolves the instant it's loaded from disk — either way the render fires right
// then. Each render produces TWO artifacts: the page PNGs (renders/output/) and
// the markdown read view (read.md) of the finished doc.
//
// Renders are funneled through a 1-slot gate (serializeRender) WITHIN this run so
// render agents don't sit idle in the agent pool waiting on Word. Correctness
// across runs is the CLI's job: `docx render --engine word` takes a cross-process
// advisory lock around the Word automation (src/core/render/engines/word-mac.ts),
// so up to ~3 concurrent workflow runs can share the single Word instance —
// their renders queue on the lock instead of corrupting each other. (The local
// corpus runner ALSO renders + reads each doc the moment it finishes, out of band,
// for during-run eyeballing; the workflow re-renders here with Word so the judge's
// PNGs are engine-consistent with the Claude arms regardless of the corpus box.)
//
// Exercise-agent tokens + wall-clock + tool split are NOT collected here — the
// skill's post-run transcript pass (scripts/exercise-metrics.ts) measures them
// accurately (the runtime gives this script no token API and bans clocks).
// ---------------------------------------------------------------------------
phase("Exercise");

// 1-slot gate that serializes this run's Word renders. Reset per run;
// serializeRender (hoisted, below) chains each render behind the previous.
let renderGate = Promise.resolve();

const pipelines = active.map((scenario) => {
	const exerciseP = usingLocal
		? Promise.resolve(
				preByKey.has(scenario.key)
					? { ...preByKey.get(scenario.key), key: scenario.key }
					: null,
			)
		: agent(exercisePrompt(scenario), {
				label: `exercise:${scenario.key}`,
				phase: "Exercise",
				model: exerciseModel,
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
					model: JUDGE_MODEL,
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
	`Exercise done: ${exercises.map((e) => `${e.key}=${e.status || e.completed}`).join(", ")}`,
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
// Phase 4 — Synthesize (fable). Prioritized improvement report.
// ---------------------------------------------------------------------------
phase("Synthesize");
const report = await agent(synthPrompt(active, exercises, verdicts), {
	label: "synthesize",
	phase: "Synthesize",
	model: SYNTH_MODEL,
	agentType: "general-purpose",
});

return { arm, report, runDir, binary, exercises, verdicts };

// ===========================================================================
// Prompt builders (hoisted function declarations)
// ===========================================================================

// Dispatch the exercise prompt by arm. The docx-cli arm (default) is the original
// prompt; the anthropic-docx-skill arm is the fair competitor analog — same task,
// same scenario folder, same structured report, but its toolset is the Anthropic
// docx skill (read its SKILL.md, run its scripts, hand-edit OOXML) instead of the
// docx-cli binary.
function exercisePrompt(scenario) {
	return arm === "anthropic-docx-skill"
		? exercisePromptAnthropic(scenario)
		: exercisePromptDocxCli(scenario);
}

function exercisePromptDocxCli(scenario) {
	const dir = `${runDir}/${scenario.key}`;
	const workLine =
		scenario.kind === "edit"
			? `Your working document (already a private copy — edit it IN PLACE, do NOT use -o/--output):\n  ${dir}/${scenario.doc}`
			: `You are authoring from scratch. Create your output at EXACTLY this path:\n  ${dir}/${scenario.doc}`;

	return `You are stress-testing **docx-cli**, a command-line tool that lets agents read, edit, and comment on Microsoft Word (.docx) files. You are playing the role of a CAPABLE-BUT-FRESH agent: you have NOT used this tool before. Discover everything you need from the tool's own help — do not assume flags.

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
  ${dir}/task.md     — what you've been asked to do: the full request, in plain terms
  ${dir}/assets/     — any additional input files (data, images). \`ls\` it; if it holds files, Read them. It may be empty.

## Your task — ${scenario.bucket}  (scenario: ${scenario.key})
Read ${dir}/task.md, then carry the task out on the working document above. The request describes the OUTCOME the person wants — it's on you to work out which of the tool's features get you there (that discovery is part of what's being measured).

## Rules
- STAY IN YOUR SCENARIO FOLDER. The only document you touch is the working file above; the only other files you read live under ${dir} (task.md, assets/). Do NOT search the wider filesystem (no roaming \`find\`, no \`ls\`/\`cat\` of other directories), and do NOT copy files in from elsewhere. The run workspace contains OTHER scenarios' folders with look-alike fixtures that are NOT yours — touching them corrupts the test and wastes calls. If something seems missing, re-read your working file; don't go hunting.
- Use ONLY the docx-cli executable above for document operations. Do NOT hand-edit the XML, unzip the .docx, or reach for any other docx library. The whole point is to test THIS tool.
- You MAY use the Read tool on your task.md / assets, and run \`${binary} read <file>\` to inspect your progress.
- Locators (p0, t0:r1c2:p0, sN, etc.) shift after structural edits — re-read when needed. Prefer batch operations where the tool offers them.
- If a command fails or confuses you, try at most ~3 reasonable alternatives, then RECORD it as friction and move on. Do not loop forever on one step.
- Make a genuine, complete attempt. Finish the task if you can.

## What to report (this is the actual product of your run)
Return the structured result. Be brutally honest — surfacing rough edges is the entire purpose. Do NOT tally your tool calls or tokens — the harness measures those from your transcript; your job is the qualitative account:
- completed: yes | partial | no
- summary: one short paragraph of what you actually accomplished.
- deadEnds: wrong turns, retries, things you expected to work but didn't (name the specific command and what it did).
- frictions: concrete "what could have been easier?" points, each with severity (blocker | major | minor) and a suggested fix. Include discoverability gaps (couldn't find the right command/flag), confusing output, and anything that made a weak agent likely to fail.
- outputPath: the absolute path to the .docx you produced (it should be ${dir}/${scenario.doc}).`;
}

// The competitor arm's exercise prompt. Deliberately MIRRORS exercisePromptDocxCli —
// same role framing, same scenario folder, same isolation rules, same structured
// report — so the only variable is the toolset. Here the toolset is the Anthropic
// docx skill at ${competitorDir}: the agent reads its SKILL.md and follows it, and is
// explicitly permitted to do everything that skill prescribes (unpack the .docx,
// hand-edit OOXML, run the Python helper scripts, use python-docx / the Node \`docx\`
// library / pandoc). Giving the competitor its full, intended toolset is the fairness
// requirement — an under-equipped competitor would void the comparison.
function exercisePromptAnthropic(scenario) {
	const dir = `${runDir}/${scenario.key}`;
	const workLine =
		scenario.kind === "edit"
			? `Your working document (already a private copy — edit it IN PLACE, inside your scenario folder):\n  ${dir}/${scenario.doc}`
			: `You are authoring from scratch. Create your output at EXACTLY this path:\n  ${dir}/${scenario.doc}`;

	return `You are stress-testing **the Anthropic "docx" Agent Skill**, the official skill for creating, reading, editing, and commenting on Microsoft Word (.docx) files. You are playing the role of a CAPABLE-BUT-FRESH agent: you have NOT used this skill before. Discover everything you need from the skill's own instructions — do not assume a workflow.

The docx skill is installed at this absolute path:
  ${competitorDir}

Start by orienting yourself — READ the skill's instructions and follow them:
  ${competitorDir}/SKILL.md     (the skill itself — read it FULLY before you start; it defines the workflow, the helper scripts, and the conventions)
The skill's helper scripts live under:
  ${competitorDir}/scripts/      (e.g. scripts/office/unpack.py, scripts/office/pack.py, scripts/comment.py, scripts/accept_changes.py — resolve the SKILL.md's relative paths against ${competitorDir})

${workLine}

Everything you need for the TASK is in YOUR scenario folder:
  ${dir}
Read these with the Read tool before you start:
  ${dir}/task.md     — what you've been asked to do: the full request, in plain terms
  ${dir}/assets/     — any additional input files (data, images). \`ls\` it; if it holds files, Read them. It may be empty.

## Your task — ${scenario.bucket}  (scenario: ${scenario.key})
Read ${dir}/task.md, then carry the task out on the working document above. The request describes the OUTCOME the person wants — it's on you to work out, FROM THE SKILL, which techniques get you there (that discovery is part of what's being measured).

## Rules
- STAY IN YOUR SCENARIO FOLDER for task work. The only document you touch is the working file above; the only task inputs you read live under ${dir} (task.md, assets/). Do all unpacking/scratch work INSIDE ${dir} (e.g. unpack into ${dir}/unpacked). Do NOT search the wider filesystem (no roaming \`find\`, no \`ls\`/\`cat\` of other directories) and do NOT copy files in from elsewhere — the run workspace holds OTHER scenarios' look-alike fixtures that are NOT yours, and touching them corrupts the test. The ONE exception: you MAY read and run the skill's own files under ${competitorDir}.
- Use the Anthropic docx skill — its documented workflow and helper scripts — to do the work. You ARE permitted (and expected, where the SKILL.md directs) to unpack the .docx, hand-edit the unpacked OOXML XML, run the skill's Python scripts, and use python-docx, the Node \`docx\` library, and pandoc. These are all already installed. This is exactly what's being tested: the skill, with the full toolset a real user of it would have.
- You MAY use the Read tool on your task.md / assets and on the skill's own files, and inspect your progress however the skill suggests (e.g. \`pandoc\`, or re-unpacking the .docx).
- If a step fails or confuses you, try at most ~3 reasonable alternatives, then RECORD it as friction and move on. Do not loop forever on one step.
- Make a genuine, complete attempt. Finish the task if you can.

## What to report (this is the actual product of your run)
Return the structured result. Be brutally honest — surfacing rough edges is the entire purpose. Do NOT tally your tool calls or tokens — the harness measures those from your transcript; your job is the qualitative account:
- completed: yes | partial | no
- summary: one short paragraph of what you actually accomplished.
- deadEnds: wrong turns, retries, things you expected to work but didn't (name the specific step and what it did).
- frictions: concrete "what could have been easier?" points, each with severity (blocker | major | minor) and a suggested fix. Include discoverability gaps, confusing output, and anything that made a weak agent likely to fail.
- outputPath: the absolute path to the .docx you produced (it should be ${dir}/${scenario.doc}).`;
}

function renderPrompt(target) {
	const baselineBlock = target.baselineDoc
		? `\n  3. render the PRISTINE BASELINE ${target.baselineDoc} into ${target.baselineOutDir} (the "before" the output is compared against)
  4. save the baseline's markdown read view: \`${binary} read ${target.baselineDoc} > ${target.baselineMdPath}\``
		: "";
	const baselineReturn = target.baselineDoc
		? ` the list of baseline page PNG paths, the \`baselineMarkdownPath\` (${target.baselineMdPath} if that read succeeded, else empty),`
		: " an empty baseline page list and empty \`baselineMarkdownPath\` (this scenario has no pristine source — it was authored from scratch),";

	return `You are the RENDER step of an evaluation harness. Produce, for the finished .docx AND (when there's a pristine source) its "before" baseline: (1) page PNGs via docx-cli's render command, driven by Microsoft **Word**, and (2) a markdown read view saved to a file. Both the OUTPUT and the BASELINE get their own PNGs + read.md so the judge can compare before/after both visually and textually.

The CLI executable:
  ${binary}

Command shapes (confirm with \`${binary} render --help\` and \`${binary} read --help\`):
  ${binary} render <FILE> --engine word --out <DIR>       # → page PNGs
  ${binary} read <FILE> > <FILE.md>                        # → markdown read view (prints to stdout; redirect to a file)

Renders are safe to run back-to-back: the CLI itself serializes Word access across processes with a lock, so a render may briefly WAIT if another run holds Word. Give each render command a generous timeout (10 minutes) and run them SEQUENTIALLY — never background one or start a second before the first returns. (\`read\` is instant and never touches Word.)

Create output directories with \`mkdir -p\` as needed. For this target (key "${target.key}"):
  1. render ${target.output} into ${target.outDir}
  2. save the output's markdown read view: \`${binary} read ${target.output} > ${target.mdPath}\` (write it even if the render step failed — the markdown does not depend on Word).${baselineBlock}

Capture the produced PNG page paths (each render command prints them). If a render fails, capture the error text and move on — do not retry more than once. If a \`read\` fails, note it in \`error\` and return an empty path for that side.

Return the structured result with ONE entry in \`scenarios\` for key "${target.key}": whether the output rendered, the list of output page PNG paths, the \`markdownPath\` (${target.mdPath} if that read succeeded, else empty),${baselineReturn} and any error text.`;
}

function judgePrompt(scenario, exercise, render) {
	const dir = `${runDir}/${scenario.key}`;
	// The rubric is read from the PRISTINE source, not the run dir — the stage step
	// strips criteria.md from the agent's workspace so the agent can't see the answer
	// key, so the staged copy has no criteria.md to read.
	const criteriaPath = `${scenariosDir}/${scenario.key}/criteria.md`;
	const outputPath =
		(exercise && exercise.outputPath) || `${dir}/${scenario.doc}`;
	const reviewPath = `${dir}/review.md`;
	// Every backend hands the judge the SAME fields — the qualitative account. Tool
	// counts are deliberately absent (they're measured post-run from transcripts or
	// the local ledger, not judged). The completion signal differs by arm: a Claude
	// exercise self-reports `completed` (yes|partial|no); a local exercise carries a
	// `status` (completed|failed = did the harness process finish or get killed). Both
	// are passed through; JSON.stringify drops whichever is undefined for this arm.
	const exerciseJson = exercise
		? JSON.stringify(
				{
					status: exercise.status,
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
	// Local-arm context: a run the harness couldn't finish (watchdog kill / crash)
	// left a partial document — tell the judge so it attributes obviously-unfinished
	// output to the run being cut short, not necessarily a tool defect.
	const statusNote =
		exercise && exercise.status === "failed"
			? `\n\n**Run status: FAILED** — the local harness run was cut short (watchdog timeout or crash) before it finished, so the document is whatever partial state it reached. Grade the ACTUAL rendered output; attribute clearly-unfinished content to the run being killed (a run/harness limitation), not automatically a docx-cli defect.`
			: "";
	const pages = (render && render.pages) || [];
	const baselinePages = (render && render.baselinePages) || [];
	const markdownPath = (render && render.markdownPath) || "";
	const baselineMarkdownPath = (render && render.baselineMarkdownPath) || "";
	const renderLine = render
		? [
				`Rendered: ${render.rendered}.`,
				`OUTPUT page PNGs:\n${pages.map((p) => `  ${p}`).join("\n") || "  (none)"}`,
				markdownPath ? `OUTPUT markdown read view: ${markdownPath}` : "",
				baselinePages.length
					? `BASELINE (pristine "before") page PNGs:\n${baselinePages.map((p) => `  ${p}`).join("\n")}`
					: "",
				baselineMarkdownPath
					? `BASELINE markdown read view: ${baselineMarkdownPath}`
					: "",
				render.error ? `Render error: ${render.error}` : "",
			]
				.filter(Boolean)
				.join("\n")
		: "(no render result for this scenario)";

	// Competitor-arm judge preamble (empty for docx-cli) — see ARMS.
	const armNote = ARMS[arm].judgeNote;

	return `You are a STRICT evaluator judging whether ${toolName} let a weak (${modelLabel}) agent complete a real task, and whether the result is correct and well-formed. Be skeptical: neither the agent's self-reported \`completed\` nor the run's \`status\` is a grade — they describe the run, not the output. Verify the actual document before you conclude anything.${armNote}${statusNote}

## Scenario: ${scenario.key} — ${scenario.bucket}
Two files define this evaluation — READ BOTH first:
  ${dir}/task.md       — the request the agent was given (plain language, what the person wanted)
  ${criteriaPath}      — the ground-truth GRADING RUBRIC: the precise, tool-specific checks that define "correct" (the agent never saw this — it was withheld from the agent's workspace; read it from this pristine path)
The criteria.md is the authority on what passes; task.md is the human request it's graded against.

## The weak agent's account of its run
${exerciseJson}

## Renders produced (Word)
${renderLine}

## How to judge
1. READ ${dir}/task.md (the request) and ${criteriaPath} (the grading rubric you judge against).
2. READ **all** of this scenario's evidence with the Read tool — every page PNG AND every markdown read view listed in "Renders produced" above.${
		hasBaseline(scenario)
			? ` That means all FOUR: the OUTPUT page PNGs, the OUTPUT markdown read view (${markdownPath || "read.md"}), the BASELINE (pristine "before") page PNGs, AND the BASELINE markdown read view (${baselineMarkdownPath || "baseline read.md"}). Read them all before you judge — do not skip the baseline.`
			: " (This scenario was authored from scratch — read the OUTPUT page PNGs and the OUTPUT markdown read view; there is no baseline.)"
	} Look at the PNGs critically — does the document accomplish the task and look right (layout, no leftover placeholders/highlights, tables intact, figure present, columns present, etc.)? The markdown read views give you the exact text/structure.
3. Run \`${binary} read ${outputPath}\` to confirm the changes SURVIVE THE WRITE→READ LOOP — this is docx-cli's core invariant; an edit that isn't retrievable on the next read is a failure. Use \`--ast\` if you need structure (e.g. section columns, tracked changes), and \`${binary} track-changes list\` / \`${binary} comments list\` where the scenario calls for them.
4. ${
		hasBaseline(scenario)
			? `BEFORE/AFTER COMPARISON — this scenario started from a pristine source, so you have BOTH a baseline and an output render (all read in step 2). Compare them TWO ways: (a) VISUAL — the BASELINE page PNGs vs the OUTPUT page PNGs: confirm ONLY the intended content changed and all other formatting/headers/footers/layout/structure is preserved; (b) TEXTUAL — diff the BASELINE markdown read view against the OUTPUT markdown read view: this is the precise record of what text/structure changed, so an unintended change that's hard to spot in the render is obvious here.`
			: "Cross-check the OUTPUT render and its markdown read view against the criteria — this scenario was authored from scratch, so there is no baseline to diff against."
	}
5. Read the agent's account (summary, deadEnds, frictions) as qualitative UX evidence — an agent that thrashed or hit dead ends on a simple task is a UX signal; cite the specific commands or steps it names. IMPORTANT: this informs the agent-struggle / UX dimension ONLY. Do NOT downgrade \`taskSuccess\`, \`rendersCorrectly\`, or \`formattingPreserved\` because the agent struggled — those are judged purely from the render + the write→read loop. A correct output reached via a painful path is still a task SUCCESS with a UX demerit. (Tool-call counts and tokens are measured from transcripts after the run — do not estimate or judge them here.)
6. Separate two questions: did the AGENT struggle (a UX problem) vs. is the TOOL broken (a bug)? Both matter.

The CLI executable for your verification commands: ${binary}

## Write your review to disk
After you've judged, WRITE a human-readable Markdown review to EXACTLY this path (use the Write tool):
  ${reviewPath}
The review must include: the scenario key + bucket, your verdict (task success, renders correctly, formatting preserved, survived read loop), a **Merits** section (what went right), a **Demerits** section (each defect with its severity and the concrete evidence you saw in the render or read output), and a **Frictions** section (the agent's reported frictions/dead-ends and your one-line read on whether the path to the result was smooth or a slog). This file is the saved judge's review for this task — make it complete and self-contained. ${judgeMetricsNote}

Then return the structured verdict. Record BOTH sides for this task:
- merits: what went right (what the tool made easy, what the agent got correct, parts of the output that are well-formed). Always list at least one if anything worked.
- defects: the demerits — concrete, evidence-backed failures (cite what you saw in the render or read output), each with a severity.`;
}

function synthPrompt(scenarios, exercises, verdicts) {
	const payload = JSON.stringify(
		{
			scenarios: scenarios.map((scenario) => ({
				key: scenario.key,
				bucket: scenario.bucket,
				kind: scenario.kind,
			})),
			// Strip the local arm's ledger extras — the synth reads the same qualitative
			// account the judge did. status (local process lifecycle) and completed
			// (Claude self-report) are both passed; the undefined one drops out.
			exercises: exercises.map((exercise) => ({
				key: exercise.key,
				status: exercise.status,
				completed: exercise.completed,
				summary: exercise.summary,
				deadEnds: exercise.deadEnds,
				frictions: exercise.frictions,
				outputPath: exercise.outputPath,
			})),
			verdicts,
		},
		null,
		2,
	);

	// Competitor-arm synth qualifier (empty for docx-cli) — see ARMS.
	const armClause = ARMS[arm].synthClause;
	return `You are writing the final report of an adversarial usability review of **${toolName}**. Weak (${modelLabel}) agents attempted ${scenarios.length} real document tasks; a stricter judge then graded each result against ground truth using Word renders and the write→read loop. Your audience is the engineer who maintains docx-cli${armClause}. The central question: **can weak agents actually use this tool to get real work done, and what should we fix first?**

Here is all the data (every weak agent's account + every judge verdict):
\`\`\`json
${payload}
\`\`\`

Write a thorough, prioritized Markdown report with these sections:

1. **Executive summary** — can weak agents use docx-cli today? Overall pass rate, the headline strengths, and the 2–3 biggest problems.
2. **Scoreboard** — a Markdown table: scenario | bucket | task success (success/partial/fail) | renders correctly | formatting preserved | top merit | top demerit.
2b. **Per-task merits & demerits** — for EVERY scenario, a short block listing its merits (what worked) and its demerits (defects/failures) from the judge verdicts. The user explicitly wants both sides for each task.
3. **Cross-cutting themes** — group findings into: Discoverability, CLI ergonomics / surface, Correctness & bugs, Formatting fidelity / preservation, Missing capabilities. Rank themes by impact. For each, give the EVIDENCE (which scenarios, specific commands, judge defects, verbatim friction quotes) and a concrete, actionable recommendation.
4. **Prioritized fixes** — a numbered top 5–8 list, highest leverage first, each tied to the evidence above and phrased as something the maintainer can act on (ideally pointing at the command/flag/output to change).
5. **What worked well** — what weak agents found easy; don't only criticize.

${synthMetricsNote}

Cite scenario keys throughout and quote agent friction verbatim where it's illuminating.

## Write your report to disk
After you've written the report, WRITE it to EXACTLY this path (use the Write tool):
  ${runDir}/REPORT.md
This is the saved run-level report the maintainer reads — the workflow itself does no file I/O, so if you don't write it, it is lost. Then ALSO return the COMPLETE report as your final message.`;
}

// Local backend only: read each scenario's pre-produced exercise.json off disk and
// return them verbatim. The corpus runner already produced these (ledger-measured);
// the workflow's JS can't read files, so an agent does it and hands the parsed
// objects back.
function loadPrompt(targets) {
	const lines = targets
		.map((target) => `  ${target.key}: ${target.path}`)
		.join("\n");
	return `You are the LOAD step of an evaluation harness. The exercise phase already ran — a local agent harness worked each document out of band and left a parsed RESULT json per scenario. READ each file below with the Read tool and return their contents.

Files (one per scenario):
${lines}

Return { exercises: [ … ] } where each array item is the EXACT parsed JSON object from one file — copy every field VERBATIM: key, status (older files may say completed instead — keep whichever is present), summary, deadEnds, frictions, outputPath (plus any extra fields the file carries, e.g. docxCommands and _local — keep them as-is). Do NOT summarize, truncate, reorder, or invent anything. If a file is missing or unreadable, OMIT that scenario rather than fabricating a result.`;
}

// Build the stage agent's prompt: run scripts/stage-scenario.ts once per active
// scenario. The script owns the logic (copy the folder, strip the judge-only
// criteria.md, verify the required inputs) and prints a one-line JSON verdict
// {key, staged, missing} — nonzero exit means something's wrong. The agent just
// executes and relays.
function stagePrompt(targets) {
	const lines = targets
		.map(
			(target) =>
				`  bun ${scriptsDir}/stage-scenario.ts ${target.srcDir} ${target.dstDir}${target.requireDoc ? ` --require-doc ${target.requireDoc}` : ""}`,
		)
		.join("\n");

	return `You are the STAGE step of an evaluation harness. Seed an isolated run workspace by running the staging script below once per scenario — it copies the scenario folder, strips the judge-only answer key (criteria.md), and verifies the required inputs landed, printing a one-line JSON verdict {key, staged, missing} (exit 0 = staged cleanly).

Run EXACTLY these commands, one at a time:
${lines}

Sources are pristine — never edit them, and do not stage anything beyond this list. If a command exits nonzero, capture its verdict's \`missing\` lines (or the error text if it produced no JSON); do not retry more than once.

Return the structured result:
- staged: the scenario keys whose verdict said staged:true.
- missing: every \`missing\` line from the verdicts, plus a "key: <error>" line for any command that failed without a verdict. Empty array if everything staged cleanly.`;
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

// 1-slot async gate serializing THIS run's Word renders so render agents don't
// pile up idle in the agent pool. (Cross-process safety is the CLI's advisory
// lock — see the pipeline comment.) Each call queues behind the previous render;
// the gate advances regardless of success/failure so one bad render can't wedge
// the rest. (`renderGate` is the per-run `let` declared next to the pipeline.)
function serializeRender(thunk) {
	const run = renderGate.then(thunk, thunk);
	renderGate = run.then(
		() => undefined,
		() => undefined,
	);
	return run;
}

// Render ONE scenario's output (plus its baseline, if it has a pristine source) into
// that scenario's result folder (<runDir>/<key>/renders/{output,baseline}/). Each of
// the two render dirs gets BOTH its page PNGs and a read.md (markdown read view of
// that doc), so the judge can diff before/after visually AND textually. A failed
// render degrades to a rendered:false record the judge can still handle rather than
// rejecting.
function renderOne(scenario) {
	const dir = `${runDir}/${scenario.key}`;
	const baseline = hasBaseline(scenario);
	const target = {
		key: scenario.key,
		output: `${dir}/${scenario.doc}`,
		outDir: `${dir}/renders/output/`,
		mdPath: `${dir}/renders/output/read.md`,
		baselineDoc: baseline
			? `${scenariosDir}/${scenario.key}/${scenario.doc}`
			: null,
		baselineOutDir: baseline ? `${dir}/renders/baseline/` : null,
		baselineMdPath: baseline ? `${dir}/renders/baseline/read.md` : null,
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
