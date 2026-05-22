import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { main } from "../../src/cli/index";
import { captureOutput, setVerboseAck } from "../../src/cli/respond";

const BINARY = join(import.meta.dir, "..", "..", "src", "index.ts");

export type CliResult = {
	stdout: string;
	stderr: string;
	exitCode: number;
	parsed?: unknown;
};

/** Top-level mutator verbs that emit a quiet success ack by default. The
 *  test harness auto-injects `--verbose` for these so existing assertions on
 *  the ack JSON keep working without per-test churn. */
const TOP_LEVEL_MUTATORS = new Set([
	"create",
	"insert",
	"edit",
	"delete",
	"replace",
]);

/** For verb-style commands (`comments add`, `track-changes accept`, …),
 *  these subcommands are mutators. Anything not in here (e.g. `comments
 *  list`, `track-changes list`, `images extract`) is a read. */
const SUBVERB_MUTATORS: Record<string, Set<string>> = {
	comments: new Set(["add", "reply", "resolve", "delete"]),
	images: new Set(["replace"]),
	hyperlinks: new Set(["add", "replace", "delete"]),
	tables: new Set([
		"insert-row",
		"delete-row",
		"insert-column",
		"delete-column",
		"set-widths",
		"merge",
		"unmerge",
		"borders",
	]),
	"track-changes": new Set(["on", "off", "accept", "reject"]),
};

function shouldInjectVerbose(args: string[]): boolean {
	const first = args[0];
	if (!first) return false;
	if (TOP_LEVEL_MUTATORS.has(first)) return true;
	const subverbs = SUBVERB_MUTATORS[first];
	if (!subverbs) return false;
	const second = args[1];
	if (second && subverbs.has(second)) return true;
	// `track-changes FILE on|off` carries the verb in positional 2 — the
	// dispatcher routes it to the toggle command. Look at args[2] for that
	// shape specifically.
	if (first === "track-changes") {
		const third = args[2];
		if (third === "on" || third === "off") return true;
	}
	return false;
}

function withInjectedVerbose(args: string[]): string[] {
	return shouldInjectVerbose(args) &&
		!args.includes("--verbose") &&
		!args.includes("-v")
		? [...args, "--verbose"]
		: args;
}

function parseAck(stdout: string): unknown {
	const trimmed = stdout.trim();
	if (trimmed.length === 0) return undefined;
	try {
		return JSON.parse(trimmed);
	} catch {
		return undefined; // not JSON (e.g. markdown / help text)
	}
}

/** Run the CLI in-process — the fast path used by the bulk of the suite. All
 * output funnels through respond.ts's sinks, which we capture here instead of
 * spawning `bun src/index.ts` per call (which dominated suite time). Still
 * exercises the real arg-parse → dispatch → command → respond path; the actual
 * process boundary (spawn, exit propagation, 64 KB stdout truncation) is
 * covered by tests/cli/binary-smoke.test.ts via {@link spawnCli}. */
export async function runCli(...args: string[]): Promise<CliResult> {
	let stdout = "";
	let stderr = "";
	captureOutput(
		(text) => {
			stdout += text;
		},
		(text) => {
			stderr += text;
		},
	);
	setVerboseAck(false); // reset the one module-global between calls
	let exitCode: number;
	try {
		exitCode = await main(["bun", "docx", ...withInjectedVerbose(args)]);
	} finally {
		captureOutput(null, null);
	}
	return { stdout, stderr, exitCode, parsed: parseAck(stdout) };
}

/** Spawn the real binary as a subprocess (`bun src/index.ts …`). Slow; use
 * only where the process boundary itself is under test (binary smoke tests).
 * Everything else should use the in-process {@link runCli}. */
export async function spawnCli(...args: string[]): Promise<CliResult> {
	const proc = Bun.spawn(["bun", BINARY, ...args], {
		stdout: "pipe",
		stderr: "pipe",
	});
	const stdout = await new Response(proc.stdout).text();
	const stderr = await new Response(proc.stderr).text();
	const exitCode = await proc.exited;
	return { stdout, stderr, exitCode, parsed: parseAck(stdout) };
}

export function tempWorkspace(label: string): string {
	return mkdtempSync(join(tmpdir(), `docx-cli-${label}-`));
}
