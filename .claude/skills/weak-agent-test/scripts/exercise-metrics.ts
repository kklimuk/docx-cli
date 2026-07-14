#!/usr/bin/env bun
/**
 * Roll up each exercise's cost/effort/correctness into one per-scenario metrics table,
 * for BOTH backends:
 *   - Claude (haiku/sonnet): reconstruct tokens + wall-clock + tool split from the
 *     workflow's agent transcripts (the runtime gives the workflow no token API and
 *     bans clocks, so this post-run pass reads the transcript jsonl instead).
 *   - Local harness: read the SAME numbers straight from each scenario's exercise.json
 *     `_local` block (ledger-measured by run-local-corpus.ts — the local model never
 *     self-reports). No transcripts exist for a local run, so pass `--local`.
 * Either way it also pulls in CORRECTNESS — the judge's task-success verdict, read
 * from <run_dir>/<key>/verdict.json (written by the skill from the workflow's return).
 *
 * The two token-source semantics line up column-for-column: fresh input (non-cache),
 * cache read (KV/prompt-cache reuse — cheap), output, and a cache-cost-weighted
 * "effective input". A local KV cache has no write premium, so its cache-write column
 * is 0; the reuse still shows up as cache read.
 *
 * Usage:
 *   exercise-metrics.ts <transcript_dir> <run_dir> <binary> [model]   # Claude
 *   exercise-metrics.ts --local <run_dir> [label]                     # local harness
 *
 * Writes <run_dir>/exercise-metrics.{md,json} (run-level), drops each scenario's row
 * into <run_dir>/<key>/metrics.json (so each per-task folder is self-contained), and
 * prints the Markdown section to stdout (the skill appends it to REPORT.md).
 */

const USAGE = `Usage:
  exercise-metrics.ts <transcript_dir> <run_dir> <binary> [model]   # Claude (from transcripts)
  exercise-metrics.ts --local <run_dir> [label]                     # local harness (from exercise.json)

Claude: [model] is the exercise-agent model substring to measure (default "haiku");
pass the workflow's args.model (e.g. "sonnet") so the matching agents are measured.
<transcript_dir> is the "Transcript dir" printed when the workflow was launched.
Local: [label] names the harness/model in the output (default "local").
Both write <run_dir>/exercise-metrics.{md,json}, drop each scenario's row into
<run_dir>/<key>/metrics.json, and print the Markdown section to stdout.`;

const FILE_TO_KEY: Record<string, string> = {
	"mnda.docx": "mnda",
	"invoice.docx": "invoice",
	"resume.docx": "resume",
	"contract.docx": "contract-markup",
	"contract-redlined.docx": "contract-finalize",
	"journal.docx": "eliot-journal",
};

function parseTs(value: unknown): Date | null {
	if (!value || typeof value !== "string") {
		return null;
	}
	// JS Date parses ISO 8601 with a trailing `Z` natively.
	const date = new Date(value);
	return Number.isNaN(date.getTime()) ? null : date;
}

function firstUserText(lines: any[]): string {
	// The agent's prompt: text of the first `user`-type message.
	for (const obj of lines) {
		if (obj?.type !== "user") {
			continue;
		}
		const content = obj?.message?.content;
		if (typeof content === "string") {
			return content;
		}
		if (Array.isArray(content)) {
			const out: string[] = [];
			for (const part of content) {
				if (part && typeof part === "object" && part.type === "text") {
					out.push(part.text ?? "");
				} else if (typeof part === "string") {
					out.push(part);
				}
			}
			if (out.length) {
				return out.join("\n");
			}
		}
	}
	return "";
}

function classifyScenario(prompt: string): string | null {
	const match = prompt.match(/\(scenario:\s*([a-z0-9-]+)\)/);
	if (match) {
		return match[1] ?? null;
	}
	// Fallback: map by the working/output filename mentioned in the prompt.
	for (const [filename, key] of Object.entries(FILE_TO_KEY)) {
		if (prompt.includes(filename)) {
			return key;
		}
	}
	return null;
}

function isDocxCall(toolName: unknown, toolInput: unknown, binary: string): boolean {
	if (toolName !== "Bash") {
		return false;
	}
	const command =
		toolInput && typeof toolInput === "object"
			? ((toolInput as any).command ?? "")
			: "";
	return (
		command.includes(binary) ||
		command.includes("dist/docx") ||
		/(^|\s)docx\s/.test(command)
	);
}

// Prompt-cache cost weights relative to a normal input token (uniform across models):
// a cache WRITE costs ~25% more, a cache READ ~90% less. Summing raw cache into the
// input count overstates a cache hit's cost ~10x, so effectiveInput() reweights it.
// A local KV cache has no write premium, so its cache-write term is always 0.
const CACHE_WRITE_MULT = 1.25;
const CACHE_READ_MULT = 0.1;

function effectiveInput(fresh: number, cacheWrite: number, cacheRead: number): number {
	// Cache-cost-weighted input, in normal-input-token equivalents (output excluded —
	// it bills at a different rate, so keep it separate).
	return Math.round(fresh + CACHE_WRITE_MULT * cacheWrite + CACHE_READ_MULT * cacheRead);
}

type TaskSuccess = "success" | "partial" | "fail" | null;

type AgentRow = {
	model: string | null;
	scenario: string | null;
	// Correctness — the judge's task-success verdict for this scenario (from
	// <run_dir>/<key>/verdict.json), or null if the run wasn't graded / no verdict.
	taskSuccess: TaskSuccess;
	inputTokens: number;
	cacheReadTokens: number;
	cacheWriteTokens: number;
	outputTokens: number;
	effectiveInputTokens: number;
	durationSec: number | null;
	docxToolCalls: number;
	otherToolCalls: number;
	totalToolCalls: number;
};

// ── Claude source: reconstruct a row from one agent transcript ──────────────
async function measureAgent(path: string, binary: string): Promise<AgentRow> {
	const lines: any[] = [];
	const text = await Bun.file(path).text();
	for (const raw of text.split("\n")) {
		const trimmed = raw.trim();
		if (!trimmed) {
			continue;
		}
		try {
			lines.push(JSON.parse(trimmed));
		} catch {
			continue;
		}
	}

	let model: string | null = null;
	// The three input flavors bill very differently, so keep them apart instead of
	// summing into one "input" number (cache reads are the cheapest but usually the
	// largest, so a naive sum overstates cost ~10x). effectiveInput() reweights them.
	let inFresh = 0;
	let cacheRead = 0;
	let cacheWrite = 0;
	let outTokens = 0;
	let docxCalls = 0;
	let otherCalls = 0;
	let firstTs: Date | null = null;
	let lastTs: Date | null = null;

	for (const obj of lines) {
		const ts = parseTs(obj?.timestamp);
		if (ts) {
			if (firstTs === null || ts < firstTs) {
				firstTs = ts;
			}
			if (lastTs === null || ts > lastTs) {
				lastTs = ts;
			}
		}
		const message = obj?.message;
		if (!message || typeof message !== "object") {
			continue;
		}
		if (message.model) {
			model = message.model;
		}
		const usage = message.usage ?? {};
		inFresh += usage.input_tokens || 0;
		cacheWrite += usage.cache_creation_input_tokens || 0;
		cacheRead += usage.cache_read_input_tokens || 0;
		outTokens += usage.output_tokens || 0;
		const content = message.content;
		if (Array.isArray(content)) {
			for (const part of content) {
				if (part && typeof part === "object" && part.type === "tool_use") {
					if (isDocxCall(part.name, part.input, binary)) {
						docxCalls += 1;
					} else {
						otherCalls += 1;
					}
				}
			}
		}
	}

	const prompt = firstUserText(lines);
	const duration =
		firstTs && lastTs
			? Math.round(((lastTs.getTime() - firstTs.getTime()) / 1000) * 10) / 10
			: null;
	return {
		model,
		scenario: classifyScenario(prompt),
		taskSuccess: null, // filled in main() from verdict.json
		inputTokens: inFresh,
		cacheReadTokens: cacheRead,
		cacheWriteTokens: cacheWrite,
		outputTokens: outTokens,
		effectiveInputTokens: effectiveInput(inFresh, cacheWrite, cacheRead),
		durationSec: duration,
		docxToolCalls: docxCalls,
		otherToolCalls: otherCalls,
		totalToolCalls: docxCalls + otherCalls,
	};
}

// ── Local source: build a row from one scenario's exercise.json `_local` block ──
async function measureLocalScenario(
	runDir: string,
	key: string,
	label: string,
): Promise<AgentRow | null> {
	let entry: any;
	try {
		entry = await Bun.file(`${runDir}/${key}/exercise.json`).json();
	} catch {
		return null;
	}
	// `_local` (current) or `_gemma` (runs produced before the rename) — same shape.
	const meta = entry._local ?? entry._gemma ?? {};
	const fresh = Number(meta.freshTokensTotal) || 0;
	const cacheRead = Number(meta.cachedTokensTotal) || 0;
	const output = Number(meta.genTokensTotal) || 0;
	const docxCalls = Array.isArray(entry.docxCommands) ? entry.docxCommands.length : 0;
	const otherCalls = Number(entry.otherToolCalls) || 0;
	// Prefer the corpus runner's total wall-clock; fall back to model compute time
	// (ttft+gen) for older files that predate wallClockSec.
	const wallClock =
		Number(meta.wallClockSec) ||
		Math.round(((Number(meta.ttftMsTotal) || 0) + (Number(meta.genMsTotal) || 0)) / 100) /
			10;
	return {
		model: label,
		scenario: key,
		taskSuccess: null, // filled in main() from verdict.json
		inputTokens: fresh,
		cacheReadTokens: cacheRead,
		cacheWriteTokens: 0, // local KV cache has no write premium
		outputTokens: output,
		effectiveInputTokens: effectiveInput(fresh, 0, cacheRead),
		durationSec: wallClock || null,
		docxToolCalls: docxCalls,
		otherToolCalls: otherCalls,
		totalToolCalls: docxCalls + otherCalls,
	};
}

// Correctness for a scenario: the judge's taskSuccess from <run_dir>/<key>/verdict.json
// (the skill writes it from the workflow's return). Missing/unreadable → null ("—").
async function readTaskSuccess(runDir: string, key: string): Promise<TaskSuccess> {
	try {
		const verdict = await Bun.file(`${runDir}/${key}/verdict.json`).json();
		const value = verdict?.taskSuccess;
		return value === "success" || value === "partial" || value === "fail"
			? value
			: null;
	} catch {
		return null;
	}
}

function fmtTokens(n: number | null): string {
	n = n || 0;
	if (n >= 1_000_000) {
		return `${(n / 1_000_000).toFixed(1)}M`;
	}
	if (n >= 1000) {
		return `${(n / 1000).toFixed(0)}k`;
	}
	return String(n);
}

function fmtSeconds(value: number | null): string {
	// "—" for missing; otherwise the number as-is (JS never keeps trailing zeros, so
	// 12.0 → "12", 12.3 → "12.3").
	return value === null ? "—" : String(value);
}

function pct(numerator: number, denominator: number): string {
	return denominator ? `${((numerator / denominator) * 100).toFixed(0)}%` : "—";
}

function pickOutlier(rows: AgentRow[], field: keyof AgentRow): [string, number] | null {
	// (key, value) of the row maximizing `field`, or null if every row is zero/empty.
	let best: [string, number] | null = null;
	for (const row of rows) {
		const value = (row[field] as number) || 0;
		if (best === null || value > best[1]) {
			best = [row.scenario || "?", value];
		}
	}
	return best && best[1] ? best : null;
}

type Totals = {
	agents: number;
	docxToolCalls: number;
	otherToolCalls: number;
	totalToolCalls: number;
	inputTokens: number;
	cacheReadTokens: number;
	cacheWriteTokens: number;
	outputTokens: number;
	effectiveInputTokens: number;
	durationSec: number;
	// Correctness rollup across scenarios (from the judge verdicts).
	graded: number;
	success: number;
	partial: number;
	fail: number;
};

function correctnessSummary(totals: Totals): string {
	if (!totals.graded) {
		return "not graded (no verdict.json)";
	}
	return `${totals.success}/${totals.graded} success · ${totals.partial} partial · ${totals.fail} fail`;
}

function renderSection(
	rows: AgentRow[],
	totals: Totals,
	modelLabel: string,
	measuredFrom: string,
	extras: { comparison?: string } = {},
): string {
	// The run-metrics section: a totals table, a per-scenario table (with a correctness
	// column), and an outliers list. Emitted to stdout (appended to REPORT.md) and
	// reused as the body of the standalone exercise-metrics.md.
	const lines: string[] = [
		`## Run metrics — ${modelLabel} (measured from ${measuredFrom})`,
		"",
	];
	if (!rows.length) {
		lines.push(
			`_No ${modelLabel} exercise results found — nothing to measure._`,
			"",
		);
		return lines.join("\n");
	}

	lines.push(
		`Per scenario for the ${rows.length} ${modelLabel} exercise(s): correctness (the ` +
			"judge's task verdict), token cost, wall-clock time, and the docx-cli-vs-other " +
			"tool split. `task` = success/partial/fail from the judge (— = not graded). " +
			"`docx share` = docx calls ÷ total calls — a low share, or many calls for a " +
			"simple task, is a friction signal. **eff in** = cache-cost-weighted input " +
			`(cache write ×${CACHE_WRITE_MULT}, cache read ×${CACHE_READ_MULT}); output is ` +
			"kept separate (different rate). See Totals for the raw cache split. Tool/token " +
			"counts are MEASURED, never self-reported.",
		"",
	);
	lines.push(
		"### Totals",
		"",
		"| metric | value |",
		"| --- | --: |",
		`| exercises | ${totals.agents} |`,
		`| correctness | ${correctnessSummary(totals)} |`,
		`| docx-cli calls | ${totals.docxToolCalls} |`,
		`| other tool calls | ${totals.otherToolCalls} |`,
		`| total tool calls | ${totals.totalToolCalls} |`,
		`| docx share | ${pct(totals.docxToolCalls, totals.totalToolCalls)} |`,
		`| fresh input tokens (non-cache) | ${fmtTokens(totals.inputTokens)} |`,
		`| cache reads (×${CACHE_READ_MULT}) | ${fmtTokens(totals.cacheReadTokens)} |`,
		`| cache writes (×${CACHE_WRITE_MULT}) | ${fmtTokens(totals.cacheWriteTokens)} |`,
		`| **effective input** (weighted) | **${fmtTokens(totals.effectiveInputTokens)}** |`,
		`| output tokens | ${fmtTokens(totals.outputTokens)} |`,
		`| wall-clock (sum) | ${fmtSeconds(totals.durationSec)} s |`,
		"",
		"### Per scenario",
		"",
		"| scenario | task | docx | other | docx share | eff in | out | time (s) |",
		"| --- | :-- | --: | --: | --: | --: | --: | --: |",
	);
	for (const row of rows) {
		lines.push(
			`| ${row.scenario || "?"} | ${row.taskSuccess ?? "—"} ` +
				`| ${row.docxToolCalls} | ${row.otherToolCalls} ` +
				`| ${pct(row.docxToolCalls, row.totalToolCalls)} ` +
				`| ${fmtTokens(row.effectiveInputTokens)} | ${fmtTokens(row.outputTokens)} ` +
				`| ${fmtSeconds(row.durationSec)} |`,
		);
	}
	lines.push(
		`| **total** | **${correctnessSummary(totals)}** ` +
			`| **${totals.docxToolCalls}** | **${totals.otherToolCalls}** ` +
			`| **${pct(totals.docxToolCalls, totals.totalToolCalls)}** ` +
			`| **${fmtTokens(totals.effectiveInputTokens)}** | **${fmtTokens(totals.outputTokens)}** ` +
			`| **${fmtSeconds(totals.durationSec)}** |`,
	);

	const outliers: string[] = [];
	const specs: [string, keyof AgentRow, (v: number) => string][] = [
		["Most docx-cli calls", "docxToolCalls", String],
		["Most non-docx (workaround) calls", "otherToolCalls", String],
		["Most output tokens", "outputTokens", fmtTokens],
		["Most effective input", "effectiveInputTokens", fmtTokens],
		["Slowest", "durationSec", (v) => `${fmtSeconds(v)} s`],
	];
	for (const [label, field, fmt] of specs) {
		const hit = pickOutlier(rows, field);
		if (hit) {
			outliers.push(`- **${label}:** \`${hit[0]}\` (${fmt(hit[1])})`);
		}
	}
	if (outliers.length) {
		lines.push("", "### Outliers", "", ...outliers);
	}

	if (extras.comparison) {
		lines.push("", extras.comparison);
	}

	lines.push("");
	return lines.join("\n");
}

/** The most-recent prior run dir (a sibling under the same parent that has a metrics
 * json for the SAME backend and sorts before this one — the run-id timestamps sort
 * lexically), or null on the first run. Subtleties:
 * - Backend-matched: a local run only compares against prior LOCAL runs and a Claude
 *   run against Claude runs — cross-backend eff-input is apples-to-oranges.
 * - Concurrent 3-at-a-time batches share a timestamp with an -rN suffix
 *   ($TS-r1/-r2/-r3); a same-batch sibling is a REPLICATE, not a prior run — skip any
 *   candidate whose -rN-stripped name matches ours.
 * - Pre-refactor Claude runs wrote haiku-metrics.json (no `backend` tag → treated as
 *   "claude"); accept both filenames so the first post-refactor run still compares. */
async function findPriorMetrics(
	runDir: string,
	backend: string,
): Promise<{ name: string; data: any } | null> {
	const normalized = runDir.replace(/\/+$/, "");
	const slash = normalized.lastIndexOf("/");
	if (slash < 0) {
		return null;
	}
	const parent = normalized.slice(0, slash);
	const current = normalized.slice(slash + 1);
	const batchOf = (name: string) => name.replace(/-r\d+$/, "");
	// Keep the parsed data on the winner so the chosen file isn't read+parsed twice
	// (the backend check already had to parse it).
	let best: { name: string; data: any } | null = null;
	for (const pattern of ["*/exercise-metrics.json", "*/haiku-metrics.json"]) {
		for await (const path of new Bun.Glob(pattern).scan({
			cwd: parent,
			absolute: false,
		})) {
			const name = path.slice(0, path.indexOf("/"));
			if (name >= current) {
				continue; // skip self and any future run
			}
			if (batchOf(name) === batchOf(current)) {
				continue; // same concurrent batch — a replicate, not a prior run
			}
			if (best !== null && name <= best.name) {
				continue; // already have a more recent candidate — no need to read this one
			}
			let data: any;
			try {
				data = JSON.parse(await Bun.file(`${parent}/${path}`).text());
			} catch {
				continue;
			}
			// Pre-refactor Claude runs have no `backend` tag → treat as "claude".
			if ((data?.backend ?? "claude") !== backend) {
				continue; // don't compare across backends
			}
			best = { name, data };
		}
	}
	return best;
}

/** A run-over-run comparison table (this run vs the previous one) with a short,
 * data-driven interpretation naming the scenarios that drove the change. Answers
 * "are we getting better or worse?" without a human diffing two JSON files — and
 * flags the single-run weak-agent variance so a swing isn't over-read as a tool
 * regression. */
function renderComparison(
	priorName: string,
	prior: any,
	rows: AgentRow[],
	totals: Totals,
): string {
	const priorRows: any[] = Array.isArray(prior?.perScenario)
		? prior.perScenario
		: [];
	const priorByKey: Record<string, any> = {};
	for (const r of priorRows) {
		if (r?.scenario) {
			priorByKey[r.scenario] = r;
		}
	}
	// Signed delta (now − before). `fmt` shapes the magnitude: fmtTokens for token
	// columns, String (default) for raw call counts.
	const signedDelta = (
		now: number,
		before: number,
		fmt: (value: number) => string = String,
	): string => {
		const difference = now - before;
		return `${difference >= 0 ? "+" : "−"}${fmt(Math.abs(difference))}`;
	};
	const delta = (now: number, before: number) => signedDelta(now, before, fmtTokens);
	const deltaCalls = (now: number, before: number) => signedDelta(now, before);

	const out: string[] = [
		`### vs previous run (\`${priorName}\`)`,
		"",
		"docx calls and effective input, this run vs last. Single-run weak-agent " +
			"metrics are high-variance (the same prompt + tool can swing 2× on agent " +
			"choices alone), so read a per-scenario swing as a lead to investigate, not " +
			"a tool-quality verdict — the reliable signal is pass-rate and per-scenario " +
			"root cause.",
		"",
		"| scenario | docx (was → now) | eff in (was → now) |",
		"| --- | --: | --: |",
	];
	// Rank scenarios by |Δ eff in| so the interpretation can name the real drivers.
	const drivers: { key: string; dEff: number }[] = [];
	for (const row of rows) {
		const key = row.scenario || "?";
		const before = priorByKey[key];
		if (!before) {
			out.push(`| ${key} | ${row.docxToolCalls} (new) | ${fmtTokens(row.effectiveInputTokens)} (new) |`);
			continue;
		}
		const dEff = row.effectiveInputTokens - (before.effectiveInputTokens || 0);
		drivers.push({ key, dEff });
		out.push(
			`| ${key} ` +
				`| ${before.docxToolCalls} → ${row.docxToolCalls} (${deltaCalls(row.docxToolCalls, before.docxToolCalls)}) ` +
				`| ${fmtTokens(before.effectiveInputTokens)} → ${fmtTokens(row.effectiveInputTokens)} (${delta(row.effectiveInputTokens, before.effectiveInputTokens)}) |`,
		);
	}
	const pt = prior?.totals;
	if (pt) {
		out.push(
			`| **total** ` +
				`| **${pt.docxToolCalls} → ${totals.docxToolCalls} (${deltaCalls(totals.docxToolCalls, pt.docxToolCalls)})** ` +
				`| **${fmtTokens(pt.effectiveInputTokens)} → ${fmtTokens(totals.effectiveInputTokens)} (${delta(totals.effectiveInputTokens, pt.effectiveInputTokens)})** |`,
		);
	}

	// Interpretation: name the 1–2 scenarios that moved eff-input the most.
	drivers.sort((a, b) => Math.abs(b.dEff) - Math.abs(a.dEff));
	const top = drivers.filter((d) => Math.abs(d.dEff) >= 50_000).slice(0, 2);
	if (top.length && pt) {
		const netDir = totals.effectiveInputTokens >= pt.effectiveInputTokens ? "up" : "down";
		const names = top
			.map((d) => `\`${d.key}\` (${d.dEff >= 0 ? "+" : "−"}${fmtTokens(Math.abs(d.dEff))})`)
			.join(" and ");
		out.push(
			"",
			`Net effective input moved **${netDir}**, driven almost entirely by ${names}` +
				" — the rest of the corpus was roughly flat. Check those scenarios' reviews " +
				"for whether the swing was real friction or agent variance before reading it " +
				"as a regression.",
		);
	}
	return out.join("\n");
}

function renderDocument(section: string, sourceDir: string, generatedAt: string): string {
	// Standalone exercise-metrics.md: a titled, dated document wrapping the section.
	return `# Exercise metrics\n\n_Generated ${generatedAt} from \`${sourceDir}\`._\n\n${section}`;
}

function localIsoSeconds(date: Date): string {
	// Local time with the UTC offset, to the second — ISO 8601 with no sub-second
	// component (e.g. 2026-06-08T21:54:30-07:00).
	const pad = (n: number) => String(n).padStart(2, "0");
	const offsetMin = -date.getTimezoneOffset();
	const sign = offsetMin >= 0 ? "+" : "-";
	const abs = Math.abs(offsetMin);
	const offset = `${sign}${pad(Math.floor(abs / 60))}:${pad(abs % 60)}`;
	return (
		`${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}` +
		`T${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}${offset}`
	);
}

// Collect the rows for a Claude run by measuring every matching agent transcript.
async function collectClaudeRows(
	transcriptDir: string,
	binary: string,
	modelFilter: string,
): Promise<AgentRow[]> {
	const paths: string[] = [];
	for await (const path of new Bun.Glob("agent-*.jsonl").scan({
		cwd: transcriptDir,
		absolute: true,
	})) {
		paths.push(path);
	}
	paths.sort();

	// Read the transcripts concurrently (each is an independent file), then keep only
	// the exercise agents (matched by model substring) — the opus/fable stage/render/
	// judge/synth agents don't count.
	const measured = await Promise.all(paths.map((path) => measureAgent(path, binary)));
	const rows = measured.filter(
		(row) => row.model && row.model.toLowerCase().includes(modelFilter),
	);
	if (rows.length === 0 && paths.length > 0) {
		process.stderr.write(
			`exercise-metrics: 0 of ${paths.length} agent transcripts matched the model filter ` +
				`"${modelFilter}" — for a non-haiku run, pass the exercise model as the 4th arg.\n`,
		);
	}
	return rows;
}

// Collect the rows for a local run by reading every scenario's exercise.json.
async function collectLocalRows(runDir: string, label: string): Promise<AgentRow[]> {
	const keys = new Set<string>();
	for await (const rel of new Bun.Glob("*/exercise.json").scan({ cwd: runDir })) {
		keys.add(rel.slice(0, rel.indexOf("/")));
	}
	// Read each scenario's exercise.json concurrently (independent files); drop any
	// that failed to parse (measureLocalScenario returns null).
	const measured = await Promise.all(
		[...keys].map((key) => measureLocalScenario(runDir, key, label)),
	);
	const rows = measured.filter((row): row is AgentRow => row !== null);
	if (!rows.length) {
		process.stderr.write(
			`exercise-metrics --local: no exercise.json files under ${runDir} — run run-local-corpus.ts first.\n`,
		);
	}
	return rows;
}

async function main(): Promise<void> {
	const argv = Bun.argv.slice(2);
	const isLocal = argv[0] === "--local";

	let rows: AgentRow[];
	let runDir: string;
	let modelLabel: string;
	let backend: string;
	let measuredFrom: string;
	let sourceDir: string;

	if (isLocal) {
		runDir = argv[1] ?? "";
		modelLabel = argv[2] || "local";
		if (!runDir) {
			console.log(USAGE);
			process.exit(2);
		}
		backend = "local";
		measuredFrom = "the harness ledger";
		sourceDir = runDir;
		rows = await collectLocalRows(runDir, modelLabel);
	} else {
		if (argv.length < 3) {
			console.log(USAGE);
			process.exit(2);
		}
		const [transcriptDir, run, binary] = argv;
		runDir = run ?? "";
		// Optional 4th arg: the exercise-agent model substring to measure (default
		// "haiku"); pass the workflow's args.model (e.g. "sonnet") so the matching
		// agents are the ones measured, not silently filtered out.
		modelLabel = (argv[3] || "haiku").toLowerCase();
		backend = "claude";
		measuredFrom = "transcripts";
		sourceDir = transcriptDir ?? "";
		rows = await collectClaudeRows(transcriptDir ?? "", binary ?? "", modelLabel);
	}

	rows.sort((a, b) => {
		const left = a.scenario || "";
		const right = b.scenario || "";
		return left < right ? -1 : left > right ? 1 : 0;
	});

	// Correctness — pull the judge's verdict for each scenario (both backends).
	for (const row of rows) {
		if (row.scenario) {
			row.taskSuccess = await readTaskSuccess(runDir, row.scenario);
		}
	}

	const total = (field: keyof AgentRow) =>
		rows.reduce((sum, row) => sum + ((row[field] as number) || 0), 0);
	const countTask = (value: TaskSuccess) =>
		rows.filter((row) => row.taskSuccess === value).length;

	const totals: Totals = {
		agents: rows.length,
		docxToolCalls: total("docxToolCalls"),
		otherToolCalls: total("otherToolCalls"),
		totalToolCalls: total("totalToolCalls"),
		inputTokens: total("inputTokens"),
		cacheReadTokens: total("cacheReadTokens"),
		cacheWriteTokens: total("cacheWriteTokens"),
		outputTokens: total("outputTokens"),
		effectiveInputTokens: total("effectiveInputTokens"),
		durationSec: Math.round(rows.reduce((sum, row) => sum + (row.durationSec || 0), 0) * 10) / 10,
		graded: rows.filter((row) => row.taskSuccess !== null).length,
		success: countTask("success"),
		partial: countTask("partial"),
		fail: countTask("fail"),
	};

	const generatedAt = localIsoSeconds(new Date());
	const prior = await findPriorMetrics(runDir, backend);
	const comparison = prior
		? renderComparison(prior.name, prior.data, rows, totals)
		: "";
	const section = renderSection(rows, totals, modelLabel, measuredFrom, { comparison });
	const document = renderDocument(section, sourceDir || "./tmp/", generatedAt);

	const outMd = `${runDir}/exercise-metrics.md`;
	const outJson = `${runDir}/exercise-metrics.json`;
	await Bun.write(outMd, document);
	await Bun.write(
		outJson,
		JSON.stringify({ backend, label: modelLabel, perScenario: rows, totals }, null, 2),
	);

	// Per-task: drop each classified scenario's measured row into its result folder,
	// so <run_dir>/<key>/ is self-contained alongside the docx, renders/, and review.md.
	for (const row of rows) {
		if (!row.scenario) {
			continue;
		}
		await Bun.write(
			`${runDir}/${row.scenario}/metrics.json`,
			JSON.stringify(row, null, 2),
		);
	}

	// stdout is appended to REPORT.md by the skill — emit the SECTION (## …), not the
	// titled standalone doc. Lead with a blank line + thematic break so it can't glue
	// onto the synthesized report's last paragraph (which has no trailing newline) and
	// turn it into a setext heading.
	console.log(`\n\n---\n\n${section}`);
	console.error(`[wrote ${outMd} and ${outJson}]`);
}

await main();
