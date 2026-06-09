import {
	type BlockRangeReference,
	type BlockReference,
	Document,
	LocatorParseError,
	LocatorResolveError,
	PkgError,
	parseLocator,
} from "@core";
import { parseArgs } from "util";

export const EXIT = {
	OK: 0,
	GENERAL_ERROR: 1,
	USAGE_ERROR: 2,
	NOT_FOUND: 3,
} as const;

export type ErrorCode =
	| "USAGE"
	| "FILE_NOT_FOUND"
	| "PART_NOT_FOUND"
	| "NOT_A_ZIP"
	| "INVALID_LOCATOR"
	| "BLOCK_NOT_FOUND"
	| "COMMENT_NOT_FOUND"
	| "IMAGE_NOT_FOUND"
	| "IMAGE_SOURCE"
	| "HYPERLINK_NOT_FOUND"
	| "TRACKED_CHANGE_NOT_FOUND"
	| "MATCH_NOT_FOUND"
	| "TRACKED_CHANGE_CONFLICT"
	| "TABLE_STRUCTURE"
	| "RENDER_ENGINE"
	| "RENDER_FAILED"
	| "UNHANDLED";

// Output sinks. Production leaves these null and writes straight to the real
// streams; the test harness redirects them to run the CLI in-process (no
// subprocess spawn). All CLI output funnels through here, so capturing these
// two captures everything.
const stdout = async (text: string) => {
	await Bun.stdout.write(text);
};
const stderr = async (text: string) => {
	await Bun.stderr.write(text);
};
const sinks = {
	stdout,
	stderr,
};

/** Redirect CLI stdout/stderr (for in-process testing). */
export function captureOutput(
	passedStdout?: ((text: string) => Promise<void>) | null,
	passedStderr?: ((text: string) => Promise<void>) | null,
): void {
	sinks.stdout = passedStdout ?? stdout;
	sinks.stderr = passedStderr ?? stderr;
}

export async function writeStdout(text: string): Promise<void> {
	await sinks.stdout(text);
}

export async function writeStderr(text: string): Promise<void> {
	await sinks.stderr(text);
}

export async function respond(payload: unknown): Promise<void> {
	await sinks.stdout(`${JSON.stringify(payload)}\n`);
}

let verboseAck = false;

/** Switch on full JSON acks for the current process. Mutating commands call
 *  this when they parse `--verbose`/`-v`. Errors always print regardless;
 *  dry-run payloads always print regardless. */
export function setVerboseAck(verbose: boolean): void {
	verboseAck = verbose;
}

/** Mutating-command success ack. `--verbose` prints the full JSON payload (the
 *  one place `ok: true` is retained); by default it prints a concise, text-first
 *  confirmation line on stdout — what changed, and where. Exit code 0 is still
 *  the machine signal, but a silent success forced (weak) agents to re-read just
 *  to confirm the command took effect, so we echo a one-liner like every other
 *  text-first command (read/find/wc). Mutators that mint a new handle use
 *  `respondMinted` instead (the handle is their confirmation). */
export async function respondAck(payload: unknown): Promise<void> {
	if (verboseAck) {
		await respond(payload);
		return;
	}
	const line = summarizeAck(payload);
	if (line) await writeStdout(`${line}\n`);
}

/** A one-line, text-first summary of a mutator ack: `<operation> <target>`,
 *  where target is the most salient identifier the payload carries (a locator,
 *  a count, an id, a table cell, …). Falls back to the operation alone. */
function summarizeAck(payload: unknown): string | null {
	if (!payload || typeof payload !== "object") return null;
	const ack = payload as Record<string, unknown>;
	const operation = typeof ack.operation === "string" ? ack.operation : null;
	if (!operation) return null;
	const target = ackTarget(ack);
	return target ? `${operation} ${target}` : operation;
}

function ackTarget(ack: Record<string, unknown>): string | null {
	const str = (value: unknown): string | null =>
		typeof value === "string" && value.length > 0 ? value : null;
	const num = (value: unknown): number | null =>
		typeof value === "number" ? value : null;
	const plural = (count: number, noun: string): string =>
		`${count} ${noun}${count === 1 ? "" : "s"}`;

	if (str(ack.locator)) return str(ack.locator);
	const editCount = num(ack.count);
	if (editCount !== null) return plural(editCount, "change");
	const replaced = num(ack.replaced);
	if (replaced !== null) return `${plural(replaced, "occurrence")} replaced`;
	if (str(ack.id)) return str(ack.id);
	if (str(ack.commentId)) return `comment ${ack.commentId}`;
	if (str(ack.hyperlinkId)) {
		const to = str(ack.to);
		return `${ack.hyperlinkId}${to ? ` → ${to}` : ""}`;
	}
	if (str(ack.imageId)) return str(ack.imageId);
	if (str(ack.table)) {
		const position = num(ack.position);
		return position !== null ? `${ack.table} (${position})` : str(ack.table);
	}
	const applied = Array.isArray(ack.applied)
		? ack.applied.length
		: num(ack.applied);
	if (applied !== null) return plural(applied, "change");
	if (str(ack.mode)) return `tracking ${ack.mode}`;
	if (Array.isArray(ack.batch)) return plural(ack.batch.length, "change");
	// Last resort: the path. Redundant for in-place edits (they win above via
	// locator/id/count), but it's the salient new thing for `create`.
	if (str(ack.path)) return str(ack.path);
	return null;
}

/** Success output for a mutator that mints a new addressable handle the agent
 *  can't reconstruct (comment/footnote/endnote/hyperlink id, inserted-block
 *  locator). Default: print the bare locator(s), one per line, so the agent can
 *  feed them straight into `--at`. With `--verbose`: print the full `ok: true`
 *  ack instead. Errors still go through `fail()`. */
export async function respondMinted(
	locators: string[],
	verbosePayload: unknown,
): Promise<void> {
	if (verboseAck) {
		await respond(verbosePayload);
		return;
	}
	if (locators.length > 0) await writeStdout(`${locators.join("\n")}\n`);
}

/** Error output. Exit code is the canonical failure signal, so no `ok` field —
 *  the nonzero exit plus the `code`/`error` keys are unambiguous. */
export async function fail(
	code: ErrorCode,
	message: string,
	hint?: string,
): Promise<number> {
	const payload: { code: ErrorCode; error: string; hint?: string } = {
		code,
		error: message,
	};
	if (hint) payload.hint = hint;
	await respond(payload);
	return exitCodeFor(code);
}

function exitCodeFor(code: ErrorCode): number {
	switch (code) {
		case "USAGE":
		case "INVALID_LOCATOR":
			return EXIT.USAGE_ERROR;
		case "FILE_NOT_FOUND":
		case "PART_NOT_FOUND":
		case "BLOCK_NOT_FOUND":
		case "COMMENT_NOT_FOUND":
		case "IMAGE_NOT_FOUND":
		case "HYPERLINK_NOT_FOUND":
		case "TRACKED_CHANGE_NOT_FOUND":
		case "MATCH_NOT_FOUND":
			return EXIT.NOT_FOUND;
		case "NOT_A_ZIP":
		case "TRACKED_CHANGE_CONFLICT":
		case "TABLE_STRUCTURE":
		case "IMAGE_SOURCE":
		case "RENDER_ENGINE":
		case "RENDER_FAILED":
		case "UNHANDLED":
			return EXIT.GENERAL_ERROR;
	}
}

export async function openOrFail(path: string): Promise<Document | number> {
	try {
		return await Document.open(path);
	} catch (err) {
		if (err instanceof PkgError) {
			if (err.code === "FILE_NOT_FOUND") {
				return await fail("FILE_NOT_FOUND", err.message);
			}
			if (err.code === "NOT_A_ZIP") return await fail("NOT_A_ZIP", err.message);
		}
		throw err;
	}
}

/** Resolve whether one mutating command should emit tracked changes. The
 *  per-command `--track` flag forces tracking on for that command regardless
 *  of the document's global `<w:trackChanges/>` setting; without the flag, the
 *  global setting decides. Every mutator (edit/insert/delete/replace, the note
 *  verbs, images delete, the tables verbs) resolves through this one helper so
 *  `--track` behaves identically everywhere. */
export function resolveTracked(
	document: Document,
	trackFlag: unknown,
): boolean {
	return Boolean(trackFlag) || document.isTrackChangesEnabled();
}

export async function resolveBlockOrFail(
	document: Document,
	locator: string,
): Promise<BlockReference | number> {
	try {
		return document.body.resolveBlock(locator);
	} catch (err) {
		if (err instanceof LocatorResolveError) {
			return await fail("BLOCK_NOT_FOUND", err.message);
		}
		throw err;
	}
}

type ParseArgsOptions = NonNullable<Parameters<typeof parseArgs>[0]>["options"];

/** The flags every in-place mutator shares: where to write (`-o`), preview-only
 *  (`--dry-run`), ack verbosity (`-v`), and `--help`. Spread into a command's
 *  `tryParseArgs` options so the four specs live in exactly one place. */
export const SAVE_FLAGS = {
	output: { type: "string", short: "o" },
	"dry-run": { type: "boolean" },
	verbose: { type: "boolean", short: "v" },
	help: { type: "boolean", short: "h" },
} as const satisfies ParseArgsOptions;

/** Wrap `parseArgs` with the boilerplate every command repeats: fix
 *  `allowPositionals: true`, catch malformed-flag errors, and translate
 *  them to `fail("USAGE", ...)`. Saves ~7 lines per command vs the inline
 *  try/catch. Returns a number (exit code) on parse failure so the caller
 *  shorts-circuits with `if (typeof parsed === "number") return parsed;`
 *  — same discriminator pattern as `openOrFail` / `resolveBlockOrFail`. */
export async function tryParseArgs(
	args: string[],
	options: ParseArgsOptions,
	help: string,
): Promise<ReturnType<typeof parseArgs> | number> {
	try {
		return parseArgs({
			args: mergeDashLeadingValues(args, options),
			allowPositionals: true,
			options,
		});
	} catch (parseError) {
		const message =
			parseError instanceof Error ? parseError.message : String(parseError);
		return await fail("USAGE", message, help);
	}
}

/** `parseArgs` rejects `--text -$500.00` as "ambiguous" because the value starts
 *  with `-` (it can't tell a negative/`-$` value from a flag). Agents hit this on
 *  money, negative numbers, etc. Pre-merge `--flag value` → `--flag=value` when
 *  `flag` is a declared string option and `value` is dash-LED but not flag-shaped
 *  (a real flag has a letter right after the dashes; `-$`, `-5`, `-.5` do not). */
function mergeDashLeadingValues(
	args: string[],
	options: ParseArgsOptions,
): string[] {
	const stringFlags = new Set(
		Object.entries(options ?? {})
			.filter(([, spec]) => spec?.type === "string")
			.map(([name]) => name),
	);
	const flagShaped = (token: string): boolean => {
		const match = token.match(/^-+(.)/);
		return match ? /[a-zA-Z]/.test(match[1] ?? "") : false;
	};
	const out: string[] = [];
	for (let index = 0; index < args.length; index++) {
		const arg = args[index];
		if (arg === undefined) continue;
		const next = args[index + 1];
		if (
			arg.startsWith("--") &&
			!arg.includes("=") &&
			stringFlags.has(arg.slice(2)) &&
			next?.startsWith("-") &&
			!flagShaped(next)
		) {
			out.push(`${arg}=${next}`);
			index++;
			continue;
		}
		out.push(arg);
	}
	return out;
}

export async function resolveBlockRangeOrFail(
	document: Document,
	locator: string,
): Promise<BlockRangeReference | number> {
	try {
		const parsed = parseLocator(locator);
		if (parsed.kind !== "blockRange") {
			return await fail("INVALID_LOCATOR", `Expected pN-pM, got ${locator}`);
		}
		return document.body.resolveBlockRange(
			parsed.startBlockId,
			parsed.endBlockId,
		);
	} catch (err) {
		if (err instanceof LocatorParseError) {
			return await fail("INVALID_LOCATOR", err.message);
		}
		if (err instanceof LocatorResolveError) {
			return await fail("BLOCK_NOT_FOUND", err.message);
		}
		throw err;
	}
}
