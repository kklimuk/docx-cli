#!/usr/bin/env bun
/**
 * Measure the Haiku exercise agents' cost/effort from a workflow run's transcripts.
 *
 * The workflow runtime gives the script no token-usage API and bans clocks, so tokens
 * and per-agent wall-clock time can't be emitted from inside the workflow. They ARE in
 * each agent's transcript jsonl (per-message `usage`, per-line `timestamp`), so this
 * post-run pass reconstructs them — accurately, and only for the Haiku exercise agents.
 *
 * Usage:
 *   haiku-metrics.ts <transcript_dir> <run_dir> <binary_path>
 *
 * <transcript_dir> is the "Transcript dir" printed when the workflow was launched
 * (…/subagents/workflows/wf_<id>). Writes <run_dir>/haiku-metrics.md and
 * <run_dir>/haiku-metrics.json (the run-level aggregate), drops each scenario's
 * measured row into <run_dir>/<key>/metrics.json (so each per-task result folder is
 * self-contained), and prints the Markdown table to stdout.
 */

const USAGE = `Usage:
  haiku-metrics.ts <transcript_dir> <run_dir> <binary_path>

<transcript_dir> is the "Transcript dir" printed when the workflow was launched
(…/subagents/workflows/wf_<id>). Writes <run_dir>/haiku-metrics.md and
<run_dir>/haiku-metrics.json, drops each scenario's row into <run_dir>/<key>/metrics.json,
and prints the Markdown table to stdout.`;

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
const CACHE_WRITE_MULT = 1.25;
const CACHE_READ_MULT = 0.1;

function effectiveInput(fresh: number, cacheWrite: number, cacheRead: number): number {
	// Cache-cost-weighted input, in normal-input-token equivalents (output excluded —
	// it bills at a different rate, so keep it separate).
	return Math.round(fresh + CACHE_WRITE_MULT * cacheWrite + CACHE_READ_MULT * cacheRead);
}

type AgentRow = {
	model: string | null;
	scenario: string | null;
	inputTokens: number;
	cacheReadTokens: number;
	cacheWriteTokens: number;
	outputTokens: number;
	effectiveInputTokens: number;
	rawTokens: number;
	durationSec: number | null;
	docxToolCalls: number;
	otherToolCalls: number;
	totalToolCalls: number;
};

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
		inputTokens: inFresh,
		cacheReadTokens: cacheRead,
		cacheWriteTokens: cacheWrite,
		outputTokens: outTokens,
		effectiveInputTokens: effectiveInput(inFresh, cacheWrite, cacheRead),
		rawTokens: inFresh + cacheRead + cacheWrite + outTokens,
		durationSec: duration,
		docxToolCalls: docxCalls,
		otherToolCalls: otherCalls,
		totalToolCalls: docxCalls + otherCalls,
	};
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
	// Mirrors Python's `:g` — drop trailing zeros (12.0 → "12", 12.3 → "12.3").
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
	rawTokens: number;
	durationSec: number;
	docxShare?: number;
};

function renderSection(rows: AgentRow[], totals: Totals): string {
	// The '## Haiku tool & cost economy' section: a totals table, a per-scenario
	// table, and an outliers list. Emitted to stdout (appended to REPORT.md) and reused
	// as the body of the standalone haiku-metrics.md.
	const lines: string[] = ["## Haiku tool & cost economy (measured from transcripts)", ""];
	if (!rows.length) {
		lines.push(
			"_No Haiku exercise agents found in the transcripts — nothing to measure._",
			"",
		);
		return lines.join("\n");
	}

	lines.push(
		`Reconstructed from the ${rows.length} Haiku exercise agent transcript(s): ` +
			"docx-cli tool calls vs everything else, token cost, and wall-clock time. " +
			"`docx share` = docx calls ÷ total calls — a low share, or many calls for a " +
			"simple task, is a friction signal. **eff in** = cache-cost-weighted input " +
			`(cache write ×${CACHE_WRITE_MULT}, cache read ×${CACHE_READ_MULT}); output is ` +
			"kept separate (different rate). See Totals for the raw cache split.",
		"",
		"### Totals",
		"",
		"| metric | value |",
		"| --- | --: |",
		`| Haiku agents | ${totals.agents} |`,
		`| docx-cli calls | ${totals.docxToolCalls} |`,
		`| other tool calls | ${totals.otherToolCalls} |`,
		`| total tool calls | ${totals.totalToolCalls} |`,
		`| docx share | ${pct(totals.docxToolCalls, totals.totalToolCalls)} |`,
		`| fresh input tokens | ${fmtTokens(totals.inputTokens)} |`,
		`| cache reads (×${CACHE_READ_MULT}) | ${fmtTokens(totals.cacheReadTokens)} |`,
		`| cache writes (×${CACHE_WRITE_MULT}) | ${fmtTokens(totals.cacheWriteTokens)} |`,
		`| **effective input** (weighted) | **${fmtTokens(totals.effectiveInputTokens)}** |`,
		`| output tokens | ${fmtTokens(totals.outputTokens)} |`,
		`| wall-clock (sum) | ${fmtSeconds(totals.durationSec)} s |`,
		"",
		"### Per scenario",
		"",
		"| scenario | docx | other | docx share | eff in | out | time (s) |",
		"| --- | --: | --: | --: | --: | --: | --: |",
	);
	for (const row of rows) {
		lines.push(
			`| ${row.scenario || "?"} ` +
				`| ${row.docxToolCalls} | ${row.otherToolCalls} ` +
				`| ${pct(row.docxToolCalls, row.totalToolCalls)} ` +
				`| ${fmtTokens(row.effectiveInputTokens)} | ${fmtTokens(row.outputTokens)} ` +
				`| ${fmtSeconds(row.durationSec)} |`,
		);
	}
	lines.push(
		`| **total** | **${totals.docxToolCalls}** | **${totals.otherToolCalls}** ` +
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

	lines.push("");
	return lines.join("\n");
}

function renderDocument(section: string, transcriptDir: string, generatedAt: string): string {
	// Standalone haiku-metrics.md: a titled, dated document wrapping the section.
	return `# Haiku metrics\n\n_Generated ${generatedAt} from \`${transcriptDir}\`._\n\n${section}`;
}

function localIsoSeconds(date: Date): string {
	// Mirrors Python's datetime.now().astimezone().isoformat(timespec="seconds"):
	// local time with the UTC offset, to the second (e.g. 2026-06-08T21:54:30-07:00).
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

async function main(): Promise<void> {
	const argv = Bun.argv.slice(2);
	if (argv.length < 3) {
		console.log(USAGE);
		process.exit(2);
	}
	const [transcriptDir, runDir, binary] = argv;

	const paths: string[] = [];
	for await (const path of new Bun.Glob("agent-*.jsonl").scan({
		cwd: transcriptDir,
		absolute: true,
	})) {
		paths.push(path);
	}
	paths.sort();

	const rows: AgentRow[] = [];
	for (const path of paths) {
		const row = await measureAgent(path, binary ?? "");
		// Haiku exercise agents only — the opus render/judge/synth agents don't count.
		if (!row.model || !row.model.includes("haiku")) {
			continue;
		}
		rows.push(row);
	}

	rows.sort((a, b) => {
		const left = a.scenario || "";
		const right = b.scenario || "";
		return left < right ? -1 : left > right ? 1 : 0;
	});

	const total = (field: keyof AgentRow) =>
		rows.reduce((sum, row) => sum + ((row[field] as number) || 0), 0);

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
		rawTokens: total("rawTokens"),
		durationSec: Math.round(rows.reduce((sum, row) => sum + (row.durationSec || 0), 0) * 10) / 10,
	};
	totals.docxShare = totals.totalToolCalls
		? Math.round((totals.docxToolCalls / totals.totalToolCalls) * 1000) / 1000
		: 0;

	const generatedAt = localIsoSeconds(new Date());
	const section = renderSection(rows, totals);
	const document = renderDocument(section, transcriptDir ?? "./tmp/", generatedAt);

	const outMd = `${runDir}/haiku-metrics.md`;
	const outJson = `${runDir}/haiku-metrics.json`;
	await Bun.write(outMd, document);
	await Bun.write(outJson, JSON.stringify({ perScenario: rows, totals }, null, 2));

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
	// titled standalone doc, so it slots in cleanly as a sub-section.
	console.log(section);
	console.error(`[wrote ${outMd} and ${outJson}]`);
}

await main();
