import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

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

export async function runCli(...args: string[]): Promise<CliResult> {
	const injected =
		shouldInjectVerbose(args) &&
		!args.includes("--verbose") &&
		!args.includes("-v")
			? [...args, "--verbose"]
			: args;
	const proc = Bun.spawn(["bun", BINARY, ...injected], {
		stdout: "pipe",
		stderr: "pipe",
	});
	const stdout = await new Response(proc.stdout).text();
	const stderr = await new Response(proc.stderr).text();
	const exitCode = await proc.exited;
	const trimmed = stdout.trim();
	let parsed: unknown;
	if (trimmed.length > 0) {
		try {
			parsed = JSON.parse(trimmed);
		} catch {
			// not JSON — leave parsed undefined
		}
	}
	return { stdout, stderr, exitCode, parsed };
}

export function tempWorkspace(label: string): string {
	return mkdtempSync(join(tmpdir(), `docx-cli-${label}-`));
}
