#!/usr/bin/env bun
/**
 * Parse a Gemma harness (Inkling) session ledger into the weak-agent EXERCISE_SCHEMA
 * shape the weak-agent-test workflow's Render / Judge / Synthesize phases consume.
 *
 * This is the ONE piece of the Gemma exercise backend that genuinely needs code. The
 * run itself is a bare `bun <harness>/src/index.ts --prompt-file … --auto` call, and
 * staging is stage-gemma.sh — neither needs wrapping. This reads the ledger that run
 * leaves under <scenarioDir>/.inkling/sessions/<id>/ and reconstructs the tool
 * economy. Gemma is NOT asked to self-report a schema (too weak) — every number here
 * is measured from the ledger:
 *   - docx-vs-other tool split + per-call outcome  ← messages.jsonl (Bash exit_code)
 *   - closing summary                              ← last assistant text
 *   - per-turn wall-clock (ttft + gen)             ← trace.jsonl
 *
 * Usage:
 *   gemma-parse-ledger.ts <scenarioDir> <binary> [doc]
 *     <scenarioDir>  the folder the harness ran in (cwd); key = its basename
 *     <binary>       absolute docx-cli path — classifies docx vs other tool calls
 *     [doc]          working document filename (default "<key>.docx") — for outputPath
 *
 * Prints the EXERCISE_SCHEMA JSON to stdout and writes it to
 * <scenarioDir>/exercise.json.
 */

import { basename } from "path";

// docx-cli verbs that MUTATE the document (vs. read-only inspection). Used only for
// the coarse `completed` heuristic — the judge is the real arbiter of success.
const MUTATION_VERBS = new Set([
	"create",
	"edit",
	"insert",
	"delete",
	"replace",
	"tables",
	"comments",
	"footnotes",
	"endnotes",
	"hyperlinks",
	"images",
	"headers",
	"footers",
	"sections",
	"styles",
	"track-changes",
	"toggle",
	"set-default-font",
]);

const [scenarioDir, binary, docArg] = Bun.argv.slice(2);
if (!scenarioDir || !binary) {
	Bun.stderr.write("Usage: gemma-parse-ledger.ts <scenarioDir> <binary> [doc]\n");
	process.exit(1);
}
const key = basename(scenarioDir);
const doc = docArg ?? `${key}.docx`;

const result = await parseLedger(scenarioDir, binary, key, doc);
await Bun.write(`${scenarioDir}/exercise.json`, `${JSON.stringify(result, null, 2)}\n`);
await Bun.write(Bun.stdout, `${JSON.stringify(result)}\n`);

async function parseLedger(
	scenarioDir: string,
	binary: string,
	key: string,
	doc: string,
): Promise<ExerciseResult> {
	const sessionDir = await newestSession(`${scenarioDir}/.inkling/sessions`);
	const result: ExerciseResult = {
		key,
		completed: "no",
		summary: "",
		docxCommands: [],
		otherToolCalls: 0,
		deadEnds: [],
		frictions: [],
		outputPath: `${scenarioDir}/${doc}`,
		_gemma: {
			sessionDir: sessionDir ?? null,
			turns: 0,
			ttftMsTotal: 0,
			genMsTotal: 0,
			promptTokensTotal: 0,
			cachedTokensTotal: 0,
			freshTokensTotal: 0,
			genTokensTotal: 0,
		},
	};
	if (!sessionDir) {
		result.summary = "(no harness session ledger found — the run produced nothing)";
		result.deadEnds.push("Harness left no session ledger; likely crashed before its first turn.");
		return result;
	}

	const messages = await readJsonl(`${sessionDir}/messages.jsonl`);
	// Map toolCallId -> exit_code, so each Bash call's outcome is exact, not guessed.
	const exitById = new Map<string, number>();
	for (const message of messages) {
		if (message?.role !== "tool" || typeof message.toolCallId !== "string") continue;
		exitById.set(message.toolCallId, toolExitCode(message.content));
	}

	let lastAssistantText = "";
	for (const message of messages) {
		if (message?.role === "assistant" && typeof message.content === "string" && message.content.trim()) {
			lastAssistantText = message.content.trim();
		}
		if (message?.role !== "assistant" || !Array.isArray(message.toolCalls)) continue;
		for (const call of message.toolCalls) {
			const command = call?.name === "Bash" ? String(call.arguments?.command ?? "") : "";
			if (!command.includes(binary)) {
				result.otherToolCalls += 1;
				continue;
			}
			const exit = call.id != null ? exitById.get(String(call.id)) : undefined;
			const outcome: "ok" | "error" = exit === 0 ? "ok" : "error";
			result.docxCommands.push({ cmd: command, outcome });
			if (outcome === "error") {
				result.deadEnds.push(`docx command exited ${exit ?? "?"}: ${command}`);
			}
		}
	}

	result.summary = lastAssistantText || "(agent produced no closing summary text)";
	result.completed = inferCompleted(binary, result.docxCommands);

	const trace = await readJsonl(`${sessionDir}/trace.jsonl`);
	result._gemma.turns = trace.length;
	// One pass over the per-turn trace, accumulating every total. NOTE the token
	// semantics: promptTokens is the TOTAL prompt each turn (cachedTokens of it rode
	// the KV cache); the tokens actually re-decoded are prompt − cached, summed as
	// freshTokensTotal (per-turn max(0, …), never a subtract-of-sums).
	for (const entry of trace) {
		const promptTokens = Number(entry?.promptTokens) || 0;
		const cachedTokens = Number(entry?.cachedTokens) || 0;
		result._gemma.ttftMsTotal += Number(entry?.ttftMs) || 0;
		result._gemma.genMsTotal += Number(entry?.genMs) || 0;
		result._gemma.promptTokensTotal += promptTokens;
		result._gemma.cachedTokensTotal += cachedTokens;
		result._gemma.freshTokensTotal += Math.max(0, promptTokens - cachedTokens);
		result._gemma.genTokensTotal += Number(entry?.genTokens) || 0;
	}
	return result;
}

function inferCompleted(
	binary: string,
	docxCommands: { cmd: string; outcome: string }[],
): "yes" | "partial" | "no" {
	// Heuristic ONLY: an error-free run with a successful mutation => yes; any successful
	// call alongside errors (or reads only) => partial; nothing succeeded => no. The judge
	// decides real success — this must never read as agent self-report.
	if (docxCommands.length === 0) return "no";
	const anyOk = docxCommands.some((command) => command.outcome === "ok");
	if (!anyOk) return "no";
	const anyError = docxCommands.some((command) => command.outcome === "error");
	const okMutation = docxCommands.some(
		(command) => command.outcome === "ok" && hasMutationVerb(binary, command.cmd),
	);
	if (okMutation && !anyError) return "yes";
	return "partial";
}

function hasMutationVerb(binary: string, command: string): boolean {
	// A single Bash line often chains several docx invocations (`docx find … && docx
	// edit …`), so check EVERY invocation's verb — not just the first after the binary
	// path — else a chained mutation is invisible to the completed heuristic.
	return command
		.split(binary)
		.slice(1)
		.some((segment) => MUTATION_VERBS.has(segment.trim().split(/\s+/)[0] ?? ""));
}

function toolExitCode(content: unknown): number {
	if (typeof content !== "string") return -1;
	try {
		const parsed = JSON.parse(content);
		return typeof parsed?.exit_code === "number" ? parsed.exit_code : -1;
	} catch {
		return -1;
	}
}

async function newestSession(sessionsRoot: string): Promise<string | null> {
	// Pick the most recent session by meta.json.createdAt (the isolated per-scenario
	// dir normally holds exactly one, but a re-run would add more).
	const glob = new Bun.Glob("*/meta.json");
	let newest: { dir: string; createdAt: number } | null = null;
	for await (const rel of glob.scan({ cwd: sessionsRoot, onlyFiles: true })) {
		const metaPath = `${sessionsRoot}/${rel}`;
		let createdAt = 0;
		try {
			createdAt = Number((await Bun.file(metaPath).json())?.createdAt) || 0;
		} catch {}
		const dir = metaPath.slice(0, metaPath.length - "/meta.json".length);
		if (!newest || createdAt > newest.createdAt) newest = { dir, createdAt };
	}
	return newest?.dir ?? null;
}

async function readJsonl(path: string): Promise<any[]> {
	const file = Bun.file(path);
	if (!(await file.exists())) return [];
	const out: any[] = [];
	for (const line of (await file.text()).split("\n")) {
		if (!line.trim()) continue;
		try {
			out.push(JSON.parse(line));
		} catch {}
	}
	return out;
}

type ExerciseResult = {
	key: string;
	completed: "yes" | "partial" | "no";
	summary: string;
	docxCommands: { cmd: string; outcome: "ok" | "error"; note?: string }[];
	otherToolCalls: number;
	deadEnds: string[];
	frictions: { issue: string; severity: string; suggestion?: string }[];
	outputPath: string;
	_gemma: {
		sessionDir: string | null;
		turns: number;
		ttftMsTotal: number;
		genMsTotal: number;
		promptTokensTotal: number;
		cachedTokensTotal: number;
		freshTokensTotal: number;
		genTokensTotal: number;
	};
};
