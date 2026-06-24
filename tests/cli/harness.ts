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
	"sections",
	// `render` is a producer, not a doc mutator, but it's silent-by-default in
	// the same way (bare page paths unless --verbose) — inject --verbose so the
	// JSON-ack assertions keep working.
	"render",
]);

/** For verb-style commands (`comments add`, `track-changes accept`, …),
 *  these subcommands are mutators. Anything not in here (e.g. `comments
 *  list`, `track-changes list`, `images extract`) is a read. */
const SUBVERB_MUTATORS: Record<string, Set<string>> = {
	comments: new Set(["add", "reply", "resolve", "delete"]),
	endnotes: new Set(["add", "edit", "delete"]),
	footers: new Set(["set", "clear"]),
	footnotes: new Set(["add", "edit", "delete"]),
	headers: new Set(["set", "clear"]),
	images: new Set(["replace", "delete"]),
	hyperlinks: new Set(["add", "replace", "delete"]),
	styles: new Set(["set-default-font"]),
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
	// `track-changes` toggle accepts both orders: `on|off FILE` (verb-first, already
	// caught by the subverbs check above via args[1]) and the legacy `FILE on|off`
	// (verb in positional 2) — catch the latter by looking at args[2].
	if (first === "track-changes") {
		const third = args[2];
		if (third === "on" || third === "off") return true;
	}
	return false;
}

/** Query verbs that now default to a text-first (non-JSON) output and accept
 *  `--json` for the structured form. The harness injects `--json` so existing
 *  assertions on the parsed JSON keep working — mirroring the `--verbose`
 *  injection for mutators. Tests of the text default use {@link spawnCli} or
 *  assert `result.stdout` directly. */
const TEXT_FIRST_QUERIES = new Set(["find", "wc", "outline"]);

function withInjectedFlags(args: string[]): string[] {
	let result =
		shouldInjectVerbose(args) &&
		!args.includes("--verbose") &&
		!args.includes("-v")
			? [...args, "--verbose"]
			: args;
	const first = result[0];
	if (first && TEXT_FIRST_QUERIES.has(first) && !result.includes("--json")) {
		result = [...result, "--json"];
	}
	return result;
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
 * covered by tests/cli/output-contract.test.ts via {@link spawnCli}. */
export async function runCli(...args: string[]): Promise<CliResult> {
	let stdout = "";
	let stderr = "";
	captureOutput(
		async (text) => {
			stdout += text;
		},
		async (text) => {
			stderr += text;
		},
	);
	setVerboseAck(false); // reset the one module-global between calls
	let exitCode: number;
	try {
		exitCode = await main(["bun", "docx", ...withInjectedFlags(args)]);
	} finally {
		captureOutput();
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

/** Like {@link spawnCli}, but feeds `input` on the subprocess's stdin — the only
 * way to exercise the `-` (stdin) ingress for `--batch` / `--code-file` /
 * `--markdown-file` / `--from`, which the in-process {@link runCli} can't feed. */
export async function spawnCliStdin(
	input: string,
	...args: string[]
): Promise<CliResult> {
	const proc = Bun.spawn(["bun", BINARY, ...args], {
		stdin: Buffer.from(input),
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
