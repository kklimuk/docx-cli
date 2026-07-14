#!/usr/bin/env bun
/**
 * Run the weak-agent EXERCISE phase for the whole corpus on the LOCAL agent harness
 * (local-first, model built in) instead of a Claude model. Serial by design — the
 * GPU is a single resource, so concurrent local runs would contend badly. Each
 * scenario:
 *   stageScenario() + local prep  →  bun <harness>/src/index.ts --cwd … --prompt-file …  →  parseLedger()
 *     →  render + read (immediate, per scenario)
 * producing <runDir>/<key>/{<doc>, exercise.json, local-run.log, read.md,
 * renders/output/*.png} plus a run-level <runDir>/corpus.log (the top-level
 * orchestration log — which scenario, what status, how long — written by the run
 * itself, so it's ALWAYS in the run dir and never depends on a stdout redirect). A
 * failed or timed-out run does NOT abort the corpus — it's
 * recorded with status "failed" (watchdog kill or crash) and the loop moves on. Each
 * exercise.json carries a `status` of "completed" (ran to its own stop) or "failed";
 * that's process lifecycle only — the workflow's judge decides task quality.
 *
 * The moment a scenario finishes, this renders its doc to page PNGs and saves its
 * markdown read view into the SAME paths the workflow uses — so a long unattended GPU
 * run produces eyeball-able artifacts as it goes, AND those become the grading renders:
 * JUDGING still happens in the workflow, but its render step is idempotent and REUSES
 * these (re-rendering only anything missing), so nothing is rendered twice. Hand the
 * runDir to the workflow with exerciseBackend: "local" and it loads each exercise.json
 * and runs the SAME Render / Judge / Synthesize pipeline the Claude arms get — identical
 * grading is what makes the local numbers comparable to Haiku's.
 *
 * Usage:
 *   run-local-corpus.ts <scenariosDir> <runDir> <binary> <harness> [--context N] [--timeout SEC] [key...]
 *     --context N    chat context size (default 131072 = ask for the max; the harness
 *                    budgeter clamps it to whatever VRAM fits. Pass 32768 on purpose
 *                    to test the 16GB-realistic small-window config.)
 *     --timeout SEC  per-scenario watchdog kill (default 1500 = 25m)
 *     key...         optional subset of scenario keys (default: all six)
 *
 * Env: LOCAL_MODEL_PATH / LOCAL_MMPROJ_PATH override the harness's built-in model
 * (forwarded as INKLING_MODEL_PATH / INKLING_MMPROJ_PATH — e.g. for a control run on
 * a different gguf). Unset, the harness runs whatever model it ships with.
 */

import { parseLedger } from "./parse-local-ledger.ts";
import { stageScenario } from "./stage-scenario.ts";

// key | doc | kind — mirrors the workflow's SCENARIOS routing manifest.
const MANIFEST = [
	{ key: "mnda", doc: "mnda.docx", kind: "edit" },
	{ key: "invoice", doc: "invoice.docx", kind: "edit" },
	{ key: "resume", doc: "resume.docx", kind: "edit" },
	{ key: "contract-markup", doc: "contract.docx", kind: "edit" },
	{ key: "contract-finalize", doc: "contract-redlined.docx", kind: "edit" },
	{ key: "eliot-journal", doc: "journal.docx", kind: "author" },
] as const;

const { scenariosDir, runDir, binary, harness, contextSize, timeoutSec, keys } =
	parseArgs(Bun.argv.slice(2));

// Tee all corpus-level progress into <runDir>/corpus.log so the orchestration log is
// ALWAYS in the run itself — never dependent on the caller redirecting stdout. (Per-
// scenario harness output still streams to <dir>/local-run.log; this is the top-level
// "which scenario, what status, how long" log.) clog/cerr (hoisted, below) write to
// BOTH the console and this sink. If a caller ALSO redirects stdout to this same path,
// drop that redirect — it's redundant now and would race this writer.
await Bun.$`mkdir -p ${runDir}`.quiet().nothrow();
const corpusLog = Bun.file(`${runDir}/corpus.log`).writer();

// Catch a wrong harness path up front, before a 25-minute watchdog wait does.
if (!(await Bun.file(`${harness}/src/index.ts`).exists())) {
	cerr(`Harness entry not found: ${harness}/src/index.ts`);
	await corpusLog.end();
	process.exit(1);
}

for (const scenario of MANIFEST) {
	if (keys.length && !keys.includes(scenario.key)) continue;
	const dir = `${runDir}/${scenario.key}`;
	clog(`=================== ${scenario.key} (${scenario.kind}, ${scenario.doc}) ===================`);

	// A failed stage skips THIS scenario, never the corpus (a long unattended GPU
	// run must survive one bad folder) — hence the catch-to-skip.
	const staged = await stageScenario(
		`${scenariosDir}/${scenario.key}`,
		dir,
		scenario.kind === "edit" ? scenario.doc : null,
	).catch((error) => ({ staged: false, missing: [String(error)] }));
	if (!staged.staged) {
		cerr(`[${scenario.key}] STAGE FAILED — skipping: ${staged.missing.join("; ")}`);
		continue;
	}
	await prepareLocalWorkspace(dir, scenario.doc, scenario.kind);

	const started = Date.now();
	const { exitCode, signalCode, timedOut } = await runHarness(dir, timeoutSec);
	const elapsedSec = Math.round((Date.now() - started) / 1000);
	if (timedOut) {
		cerr(`[${scenario.key}] TIMEOUT at ${timeoutSec}s — watchdog killed the run`);
	}

	// status = process lifecycle only (completed vs failed); the judge owns quality.
	const { status, reason } = deriveStatus(timedOut, signalCode);
	try {
		const result = await parseLedger(dir, binary, scenario.key, scenario.doc);
		// The corpus runner is the authority on lifecycle — it saw the process die.
		// A killed/crashed run is "failed" even if it left a partial ledger; only keep
		// parseLedger's "failed" (no-ledger) when the process itself exited cleanly.
		if (status === "failed") result.status = "failed";
		// It also timed the process, so it's the authority on total wall-clock (the
		// ledger only has model compute time). This is what exercise-metrics reads.
		result._local.wallClockSec = elapsedSec;
		await Bun.write(`${dir}/exercise.json`, `${JSON.stringify(result, null, 2)}\n`);
	} catch (error) {
		cerr(`[${scenario.key}] PARSE FAILED: ${error}`);
	}

	// Render + read IMMEDIATELY so the run produces eyeball-able artifacts as it goes:
	// the OUTPUT doc, plus its pristine BASELINE source when there is one (every edit
	// scenario — author scenarios have no source). Best-effort: a render/read failure
	// never aborts the corpus (the workflow's idempotent render fills any gap).
	const baselineSrc =
		scenario.kind === "edit"
			? `${scenariosDir}/${scenario.key}/${scenario.doc}`
			: null;
	await renderAndRead(dir, scenario.doc, scenario.key, baselineSrc);

	clog(
		`[${scenario.key}] status=${status}${reason ? ` (${reason})` : ""} rc=${exitCode} elapsed=${elapsedSec}s`,
	);
}
clog("=================== corpus done ===================");
clog(`Next: run the weak-agent-test workflow with { runDir: "${runDir}", exerciseBackend: "local" } to render/judge/synthesize these results.`);
await corpusLog.end();

// Prep the scenario folder for a harness run: pre-authorize the docx binary and
// render the static prompt template. No model involved.
async function prepareLocalWorkspace(
	dir: string,
	doc: string,
	kind: "edit" | "author",
): Promise<void> {
	// Pre-authorize the docx binary so the local agent drives it unimpeded (clean
	// tool-economy signal, not the --auto permission judge); deny web/agent so it
	// can't wander.
	// NOTE: this JSON mirrors the local harness's `permissions` rule format
	// (allow/deny/ask arrays of `Tool(matcher)` rule strings; see the harness repo's
	// src/permissions). If that grammar ever changes, the harness silently IGNORES
	// rules it can't parse — update this block to match rather than debugging a
	// "why is the agent getting prompted" mystery.
	await Bun.write(
		`${dir}/.inkling/settings.json`,
		`${JSON.stringify(
			{
				permissions: {
					allow: [
						`Bash(${binary}:*)`,
						"Bash(ls:*)",
						"Bash(cat:*)",
						"Bash(pwd)",
						"Bash(mkdir:*)",
					],
					deny: ["Bash(sudo:*)", "WebSearch", "WebFetch", "Agent", "AgentPool"],
					ask: [],
				},
			},
			null,
			2,
		)}\n`,
	);

	// Render the prompt template with this scenario's binary path, working doc, and
	// the task brief INLINED ({{TASK}} = task.md's content — the small local model
	// reads the request in the prompt like a human typing the ask, instead of
	// spending a turn on `cat task.md`).
	const workline =
		kind === "author"
			? "You are authoring from scratch. Create your output at EXACTLY this path:"
			: "Your working document (already a private copy — edit it IN PLACE, do NOT pass -o/--output):";
	const template = await Bun.file(
		`${import.meta.dir}/local-exercise-prompt.md`,
	).text();
	const task = (await Bun.file(`${dir}/task.md`).text()).trim();
	await Bun.write(
		`${dir}/exercise-prompt.md`,
		template
			.replaceAll("{{WORKLINE}}", workline)
			.replaceAll("{{BINARY}}", binary)
			.replaceAll("{{DOC}}", doc)
			.replaceAll("{{TASK}}", task),
	);
}

// Render + read each scenario's artifacts the moment it finishes: the OUTPUT doc into
// <dir>/renders/output/ and, when a pristine source exists, the BASELINE into
// <dir>/renders/baseline/. Each render dir gets BOTH its page PNGs and a read.md,
// mirroring the workflow's layout EXACTLY (same paths). These are the primary renders
// for the local backend — the workflow's render step is idempotent and REUSES them
// (re-rendering only what's missing), so this isn't throwaway "eyeballing" work; it's
// what the judge grades. All best-effort (a failure logs and returns; the corpus must
// not abort, and the workflow fills any gap).
async function renderAndRead(
	dir: string,
	doc: string,
	key: string,
	baselineSrc: string | null,
): Promise<void> {
	await renderReadOne(`${dir}/${doc}`, `${dir}/renders/output`, key, "output");
	if (baselineSrc) {
		await renderReadOne(baselineSrc, `${dir}/renders/baseline`, key, "baseline");
	}
}

// Render one doc to <outDir>/*.png and save its markdown read view to <outDir>/read.md.
// The read view is engine-independent (AST-as-markdown, no rendering); the PNGs use the
// default engine — Word on the mac the local harness runs on, which is the same engine
// the workflow grades on, so the workflow reuses these instead of re-rendering.
async function renderReadOne(
	srcDoc: string,
	outDir: string,
	key: string,
	label: string,
): Promise<void> {
	if (!(await Bun.file(srcDoc).exists())) {
		cerr(`[${key}] no ${label} doc at ${srcDoc} (nothing to render/read)`);
		return;
	}
	await Bun.$`mkdir -p ${outDir}`.quiet().nothrow();
	const read = await Bun.$`${binary} read ${srcDoc}`.quiet().nothrow();
	if (read.exitCode === 0) {
		await Bun.write(`${outDir}/read.md`, read.stdout);
	} else {
		cerr(
			`[${key}] ${label} READ FAILED (continuing): ${read.stderr.toString().trim() || "(no stderr)"}`,
		);
	}
	const render = await Bun.$`${binary} render ${srcDoc} --out ${outDir}`
		.quiet()
		.nothrow();
	if (render.exitCode !== 0) {
		cerr(
			`[${key}] ${label} RENDER FAILED (continuing): ${render.stderr.toString().trim() || "(no stderr)"}`,
		);
	}
}

// Spawn the harness against the staged folder, streaming stdout+stderr to
// <dir>/local-run.log (tail -f-able during the run), with a watchdog that SIGTERMs
// a run past the ceiling.
async function runHarness(
	dir: string,
	capSec: number,
): Promise<{ exitCode: number; signalCode: string | null; timedOut: boolean }> {
	const env: Record<string, string | undefined> = { ...Bun.env };
	if (Bun.env.LOCAL_MODEL_PATH) env.INKLING_MODEL_PATH = Bun.env.LOCAL_MODEL_PATH;
	if (Bun.env.LOCAL_MMPROJ_PATH)
		env.INKLING_MMPROJ_PATH = Bun.env.LOCAL_MMPROJ_PATH;

	const logSink = Bun.file(`${dir}/local-run.log`).writer();
	const proc = Bun.spawn(
		[
			"bun",
			`${harness}/src/index.ts`,
			"--cwd",
			dir,
			"--prompt-file",
			`${dir}/exercise-prompt.md`,
			"--auto",
			"--context-size",
			String(contextSize),
		],
		{ stdout: "pipe", stderr: "pipe", env },
	);

	let timedOut = false;
	const watchdog = setTimeout(() => {
		timedOut = true;
		proc.kill("SIGTERM");
	}, capSec * 1000);
	const pump = async (stream: ReadableStream<Uint8Array> | null) => {
		if (!stream) return;
		for await (const chunk of stream) {
			logSink.write(chunk);
			logSink.flush();
		}
	};
	const [, , exitCode] = await Promise.all([
		pump(proc.stdout),
		pump(proc.stderr),
		proc.exited,
	]);
	clearTimeout(watchdog);
	await logSink.end();
	// signalCode is the signal that killed the process (e.g. "SIGSEGV" on a crash,
	// "SIGTERM" from our watchdog), or null on a clean exit — the crash/kill tell.
	return { exitCode, signalCode: proc.signalCode, timedOut };
}

// Classify the harness PROCESS outcome (lifecycle, not task quality):
//   - watchdog fired (timedOut)                → failed, reason "watchdog timeout"
//   - died on a signal we didn't send (crash)  → failed, reason "crash (SIGNAL)"
//   - exited on its own (any exit code)        → completed (parseLedger still marks
//                                                 "failed" if it left NO ledger)
// A clean nonzero exit is "completed" — the process ran to its own stop; the judge,
// not this heuristic, decides whether the resulting document is any good.
function deriveStatus(
	timedOut: boolean,
	signalCode: string | null,
): { status: "completed" | "failed"; reason: string } {
	if (timedOut) return { status: "failed", reason: "watchdog timeout" };
	if (signalCode) return { status: "failed", reason: `crash (${signalCode})` };
	return { status: "completed", reason: "" };
}

// Tee a line to the console AND <runDir>/corpus.log (flushed per line so a crash mid-
// run still leaves the log on disk). clog → stdout; cerr → stderr. Both reference the
// module-level corpusLog sink, opened once runDir is known.
function clog(message: string): void {
	console.log(message);
	corpusLog.write(`${message}\n`);
	corpusLog.flush();
}

function cerr(message: string): void {
	console.error(message);
	corpusLog.write(`${message}\n`);
	corpusLog.flush();
}

function parseArgs(argv: string[]): {
	scenariosDir: string;
	runDir: string;
	binary: string;
	harness: string;
	contextSize: number;
	timeoutSec: number;
	keys: string[];
} {
	const positional: string[] = [];
	let contextSize = 131072;
	let timeoutSec = 1500;
	// A malformed flag value must fail LOUDLY — silently falling back to the default
	// would run a 25-minute-per-scenario corpus with the wrong config.
	const numericFlag = (flag: string, raw: string | undefined): number => {
		const value = Number(raw);
		if (!raw || !Number.isFinite(value) || value <= 0) {
			console.error(`Invalid ${flag} value: ${JSON.stringify(raw)} — expected a positive number.`);
			process.exit(2);
		}
		return value;
	};
	for (let index = 0; index < argv.length; index += 1) {
		const arg = argv[index] ?? "";
		if (arg === "--context") {
			contextSize = numericFlag("--context", argv[++index]);
			continue;
		}
		if (arg === "--timeout") {
			timeoutSec = numericFlag("--timeout", argv[++index]);
			continue;
		}
		positional.push(arg);
	}
	const [scenariosDir, runDir, binary, harness, ...keys] = positional;
	if (!scenariosDir || !runDir || !binary || !harness) {
		console.error(
			"Usage: run-local-corpus.ts <scenariosDir> <runDir> <binary> <harness> [--context N] [--timeout SEC] [key...]",
		);
		process.exit(2);
	}
	return { scenariosDir, runDir, binary, harness, contextSize, timeoutSec, keys };
}
