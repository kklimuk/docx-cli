#!/usr/bin/env bun
/**
 * Parse a local-harness session ledger into the weak-agent EXERCISE shape the
 * weak-agent-test workflow's Render / Judge / Synthesize phases consume.
 *
 * This is the ONE piece of the local exercise backend that genuinely needs code. The
 * run itself is a bare `bun <harness>/src/index.ts --prompt-file … --auto` call
 * (driven by run-local-corpus.ts, which imports parseLedger()). This reads the ledger
 * that run leaves under <scenarioDir>/.inkling/sessions/<id>/ and reconstructs the
 * tool economy. The local model is NOT asked to self-report anything (too weak — and
 * the Claude arms don't self-report either) — every number here is measured from the
 * ledger:
 *   - docx-vs-other tool split + per-call outcome  ← messages.jsonl (Bash exit_code)
 *   - closing summary                              ← last assistant text
 *   - per-turn wall-clock (ttft + gen) + tokens    ← trace.jsonl
 *
 * Usage:
 *   parse-local-ledger.ts <scenarioDir> <binary> [doc]
 *     <scenarioDir>  the folder the harness ran in (cwd); key = its basename
 *     <binary>       absolute docx-cli path — classifies docx vs other tool calls
 *     [doc]          working document filename (default "<key>.docx") — for outputPath
 *
 * Prints the exercise JSON to stdout and writes it to <scenarioDir>/exercise.json.
 */

import { existsSync } from "node:fs";
import { basename } from "node:path";

if (import.meta.main) {
	const [scenarioDir, binary, docArg] = Bun.argv.slice(2);
	if (!scenarioDir || !binary) {
		await Bun.write(
			Bun.stderr,
			"Usage: parse-local-ledger.ts <scenarioDir> <binary> [doc]\n",
		);
		process.exit(1);
	}
	const key = basename(scenarioDir.replace(/\/+$/, ""));
	const doc = docArg ?? `${key}.docx`;
	const result = await parseLedger(scenarioDir, binary, key, doc);
	await Bun.write(
		`${scenarioDir}/exercise.json`,
		`${JSON.stringify(result, null, 2)}\n`,
	);
	await Bun.write(Bun.stdout, `${JSON.stringify(result)}\n`);
}

export async function parseLedger(
	scenarioDir: string,
	binary: string,
	key: string,
	doc: string,
): Promise<ExerciseResult> {
	const sessionDir = await newestSession(`${scenarioDir}/.inkling/sessions`);
	const result: ExerciseResult = {
		key,
		// Process lifecycle, NOT task quality (the judge owns quality). Default
		// "completed" — a ledger means the harness ran ≥1 turn; the no-ledger branch
		// below flips it to "failed", and the corpus runner overrides it to "failed"
		// when the watchdog killed the run or it died on a signal (see deriveStatus).
		status: "completed",
		summary: "",
		docxCommands: [],
		otherToolCalls: 0,
		deadEnds: [],
		frictions: [],
		outputPath: `${scenarioDir}/${doc}`,
		_local: {
			sessionDir: sessionDir ?? null,
			turns: 0,
			// Total process wall-clock, set by the corpus runner (it timed the run);
			// ttft+gen below is model-compute-only and excludes tool-exec time.
			wallClockSec: 0,
			ttftMsTotal: 0,
			genMsTotal: 0,
			promptTokensTotal: 0,
			cachedTokensTotal: 0,
			freshTokensTotal: 0,
			genTokensTotal: 0,
			roundTimeouts: 0,
		},
	};
	if (!sessionDir) {
		// No ledger at all = the harness died before writing its first turn (crash on
		// startup, bad model path, etc.) — that's a failed run, not a completed one.
		result.status = "failed";
		result.summary =
			"(no harness session ledger found — the run produced nothing)";
		result.deadEnds.push(
			"Harness left no session ledger; likely crashed before its first turn.",
		);
		return result;
	}

	const messages = await readJsonl(`${sessionDir}/messages.jsonl`);
	// Map toolCallId -> exit_code, so each Bash call's outcome is exact, not guessed.
	const exitById = new Map<string, number>();
	for (const message of messages) {
		if (message?.role !== "tool" || typeof message.toolCallId !== "string")
			continue;
		exitById.set(message.toolCallId, toolExitCode(message.content));
	}

	let lastAssistantText = "";
	for (const message of messages) {
		if (
			message?.role === "assistant" &&
			typeof message.content === "string" &&
			message.content.trim()
		) {
			lastAssistantText = message.content.trim();
		}
		if (message?.role !== "assistant" || !Array.isArray(message.toolCalls))
			continue;
		for (const call of message.toolCalls) {
			const command =
				call?.name === "Bash" ? String(call.arguments?.command ?? "") : "";
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

	result.summary =
		lastAssistantText || "(agent produced no closing summary text)";
	// status stays "completed" here — a ledger exists, so the harness ran to its own
	// stop. The corpus runner flips it to "failed" if the watchdog/signal killed it.

	const trace = await readJsonl(`${sessionDir}/trace.jsonl`);
	result._local.turns = trace.length;
	// One pass over the per-turn trace, accumulating every total. NOTE the token
	// semantics: promptTokens is the TOTAL prompt each turn (cachedTokens of it rode
	// the KV cache); the tokens actually re-decoded are prompt − cached, summed as
	// freshTokensTotal (per-turn max(0, …), never a subtract-of-sums).
	for (const entry of trace) {
		const promptTokens = Number(entry?.promptTokens) || 0;
		const cachedTokens = Number(entry?.cachedTokens) || 0;
		result._local.ttftMsTotal += Number(entry?.ttftMs) || 0;
		result._local.genMsTotal += Number(entry?.genMs) || 0;
		result._local.promptTokensTotal += promptTokens;
		result._local.cachedTokensTotal += cachedTokens;
		result._local.freshTokensTotal += Math.max(0, promptTokens - cachedTokens);
		result._local.genTokensTotal += Number(entry?.genTokens) || 0;
		if (entry?.roundTimeout != null && typeof entry.roundTimeout === "object") {
			result._local.roundTimeouts += 1;
		}
	}
	return result;
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
	// A harness that crashed before its first turn leaves NO sessions dir at all
	// (.inkling/ exists — the corpus runner writes settings.json into it — but
	// sessions/ doesn't). Bun.Glob.scan THROWS ENOENT on a missing cwd, so guard
	// here or the graceful "no ledger" degrade path upstream is unreachable.
	if (!existsSync(sessionsRoot)) return null;
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

export type ExerciseResult = {
	key: string;
	// Whether the harness PROCESS ran to completion ("completed") or was killed —
	// watchdog timeout or a crash signal ("failed"). This is lifecycle, NOT task
	// quality: a "completed" run can still have produced a bad document, and the
	// judge is the sole arbiter of task success. Set by the corpus runner from the
	// process outcome (see run-local-corpus.ts deriveStatus).
	status: "completed" | "failed";
	summary: string;
	docxCommands: { cmd: string; outcome: "ok" | "error"; note?: string }[];
	otherToolCalls: number;
	deadEnds: string[];
	frictions: { issue: string; severity: string; suggestion?: string }[];
	outputPath: string;
	_local: {
		sessionDir: string | null;
		turns: number;
		// Total process wall-clock in seconds (corpus runner stamps this; 0 until then).
		// Includes tool-exec time, so it's the number comparable to a Claude agent's
		// first→last timestamp span — unlike ttft+gen, which is model compute only.
		wallClockSec: number;
		ttftMsTotal: number;
		genMsTotal: number;
		promptTokensTotal: number;
		cachedTokensTotal: number;
		freshTokensTotal: number;
		genTokensTotal: number;
		// Per-round wall-clock breaker (harness): how many rounds were aborted for
		// running past the per-round timeout — a ballooning-generation signal.
		roundTimeouts: number;
	};
};
