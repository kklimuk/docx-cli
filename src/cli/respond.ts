import {
	type BlockRangeReference,
	type BlockReference,
	Document,
	LocatorParseError,
	LocatorResolveError,
	PkgError,
	parseLocator,
} from "@core";
import { Pkg } from "@core/ast/document/package";
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
	| "RELATIONSHIP_NOT_FOUND"
	| "TRACKED_CHANGE_NOT_FOUND"
	| "MATCH_NOT_FOUND"
	| "TRACKED_CHANGE_CONFLICT"
	| "TABLE_STRUCTURE"
	| "INVALID_XML"
	| "VALIDATION_FAILED"
	| "RENDER_ENGINE"
	| "RENDER_FAILED"
	| "UPGRADE_FAILED"
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

/** The `--dry-run` preview shared by the in-place edit verbs (`edit`, `code
 *  edit`, `equations edit`, `tasks check`/`uncheck`): an in-place edit shifts no
 *  ids, so the preview just echoes the operation + locator (and `--output` when
 *  set). Always prints regardless of `--verbose`. */
export async function respondEditDryRun(
	filePath: string,
	locator: string,
	outputPath: string | undefined,
): Promise<number> {
	await respond({
		operation: "edit",
		dryRun: true,
		path: filePath,
		locator,
		...(outputPath ? { output: outputPath } : {}),
	});
	return EXIT.OK;
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
export async function respondAck(
	payload: unknown,
	layoutHint?: string,
): Promise<void> {
	if (verboseAck) {
		await respond(payload);
		return;
	}
	const parts = [summarizeAck(payload), layoutHint].filter(
		(part): part is string => Boolean(part),
	);
	if (parts.length > 0) await writeStdout(`${parts.join("\n")}\n`);
}

/** A render-verify nudge for layout-affecting mutators. `read` shows text and
 * structure but NOT how the page actually looks (multi-column flow, cell-width
 * wraps, image sizing, where content lands), so a weak agent that trusts a
 * clean `read` after a layout edit ships a broken render. Layout mutators
 * (`tables set-widths`, `columns`, `insert --section/--page-break/--image/
 * --table`) pass this to `respondAck`/`respondMinted` so the reminder lands at
 * the moment of success, not just in `--help`. */
export function renderVerifyHint(path: string): string {
	return `↳ layout changed — verify it renders right: docx render ${path} --out pages/ (read shows text/structure, NOT page layout: columns, wraps, image sizing)`;
}

/** The canonical render-and-look example line the layout-affecting HELPs embed
 * (each above its own command-specific lead-in). One wording everywhere, so a
 * weak agent recognizes the identical nudge across commands. The ack-side twin
 * is `renderVerifyHint` above. */
export const RENDER_VERIFY_EXAMPLE = `  docx render doc.docx --out pages/     # writes pages/page-001.png, … — read them`;

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
	if (str(ack.font)) return str(ack.font);
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
	layoutHint?: string,
): Promise<void> {
	if (verboseAck) {
		await respond(verbosePayload);
		return;
	}
	const parts = [...locators];
	if (layoutHint) parts.push(layoutHint);
	if (parts.length > 0) await writeStdout(`${parts.join("\n")}\n`);
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

/** A per-entry `--batch` validation failure. Carries the `fail()` triple so a
 *  batch reader can validate every entry inside one try/catch and translate the
 *  throw in a single line — shared by every `--batch` surface so they can't
 *  drift on the error contract. */
export class EntryError extends Error {
	constructor(
		public code: ErrorCode,
		message: string,
		public hint?: string,
	) {
		super(message);
		this.name = "EntryError";
	}
}

function exitCodeFor(code: ErrorCode): number {
	switch (code) {
		case "USAGE":
		case "INVALID_LOCATOR":
		case "INVALID_XML":
			return EXIT.USAGE_ERROR;
		case "FILE_NOT_FOUND":
		case "PART_NOT_FOUND":
		case "BLOCK_NOT_FOUND":
		case "COMMENT_NOT_FOUND":
		case "IMAGE_NOT_FOUND":
		case "HYPERLINK_NOT_FOUND":
		case "RELATIONSHIP_NOT_FOUND":
		case "TRACKED_CHANGE_NOT_FOUND":
		case "MATCH_NOT_FOUND":
			return EXIT.NOT_FOUND;
		case "NOT_A_ZIP":
		case "TRACKED_CHANGE_CONFLICT":
		case "TABLE_STRUCTURE":
		case "VALIDATION_FAILED":
		case "IMAGE_SOURCE":
		case "RENDER_ENGINE":
		case "RENDER_FAILED":
		case "UPGRADE_FAILED":
		case "UNHANDLED":
			return EXIT.GENERAL_ERROR;
	}
}

export async function openOrFail(path: string): Promise<Document | number> {
	try {
		return await Document.open(path);
	} catch (err) {
		return pkgOpenError(err);
	}
}

/** Open just the OPC package — no `Document`, so no body-AST build. For verbs
 *  that only touch parts (`docx validate`, `raw part list`/`get`); building the
 *  body would be pure waste. Maps `PkgError` to the CLI error shape exactly as
 *  {@link openOrFail} does. */
export async function openPkgOrFail(path: string): Promise<Pkg | number> {
	try {
		return await Pkg.open(path);
	} catch (err) {
		return pkgOpenError(err);
	}
}

function pkgOpenError(err: unknown): Promise<number> {
	if (err instanceof PkgError) {
		if (err.code === "FILE_NOT_FOUND") {
			return fail("FILE_NOT_FOUND", err.message);
		}
		if (err.code === "NOT_A_ZIP") return fail("NOT_A_ZIP", err.message);
	}
	throw err;
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

/** Weak agents hold a locator from an earlier read, run a single-shot structural
 *  mutation (insert/delete), then fire a second single-shot command at the now-stale
 *  id. The nonzero exit is the ONE moment they reliably read output, so the recovery
 *  belongs on the error, not only in `--help` (which they act before reading). Worded
 *  to match the codebase's standing line — ids shift after STRUCTURAL edits, NOT the
 *  in-place content/format edit that leaves the edited locator unchanged. */
const STALE_LOCATOR_HINT =
	"Locator ids are positional and shift after structural edits (insert/delete/section changes) — an id from an earlier read may now point elsewhere or be gone. Re-read the file to get current ids, then retry. Applying several changes? Do them from ONE read with --batch (edit/insert/delete/replace/comments) so ids never go stale mid-run.";

export async function resolveBlockOrFail(
	document: Document,
	locator: string,
): Promise<BlockReference | number> {
	try {
		return document.body.resolveBlock(locator);
	} catch (err) {
		if (err instanceof LocatorResolveError) {
			return await fail("BLOCK_NOT_FOUND", err.message, STALE_LOCATOR_HINT);
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
	// `--help`/`-h` anywhere in flag position (before a bare `--`) wins over
	// everything — an agent reaching for help must never hit a parse error first
	// (`replace --batch --help` used to die "ambiguous"). parseArgs never consumes a
	// bare `--help` as a flag value, so this can't steal an intended value; a literal
	// "--help" value still travels via `--flag=--help`, `--text-file`, or `--batch`.
	// Help is handled here for every command — a command's own `if (values.help)`
	// check is now redundant (harmless; it just never fires).
	if (hasHelpFlag(args)) {
		await writeStdout(help);
		return EXIT.OK;
	}
	const { args: normalized, placeholders } = normalizeDashLeading(
		args,
		options,
	);
	try {
		const parsed = parseArgs({
			args: normalized,
			allowPositionals: true,
			options,
		});
		if (placeholders.size === 0) return parsed;
		// Restore the dash-led positionals we shielded from parseArgs (below).
		return {
			...parsed,
			positionals: parsed.positionals.map(
				(value) => placeholders.get(value) ?? value,
			),
		};
	} catch (parseError) {
		const message =
			parseError instanceof Error ? parseError.message : String(parseError);
		return await fail("USAGE", message, help);
	}
}

/** `--help`/`-h` present as a standalone token before any bare `--` separator? */
function hasHelpFlag(args: string[]): boolean {
	const doubleDash = args.indexOf("--");
	const scan = doubleDash === -1 ? args : args.slice(0, doubleDash);
	return scan.includes("--help") || scan.includes("-h");
}

/** One pre-parse pass that makes dash-led values survivable. `parseArgs` rejects
 *  any `-`-leading token it can't map to a flag, which bites two ways: a flag's
 *  VALUE (`--text -$500.00`) and a bare POSITIONAL (`find FILE -$500.00`,
 *  `replace FILE -old -new`). Agents hit both on money/negatives, and quoting can't
 *  fix it (it's arg parsing, not the shell). We:
 *   1. merge a long string flag with a dash-led value → `--flag=value`;
 *   2. shield a dash-led POSITIONAL (a token that's neither a flag nor a flag's
 *      value) behind an internal sentinel, restored in `tryParseArgs` after the
 *      parse.
 *  A token is "flag-shaped" (a real flag, left untouched) only when a LETTER follows
 *  the dashes — `-$`, `-5`, `-.5`, `-00.00` are not, so a real flag typo like `-all`
 *  still errors. Prepending a bare `--` would be simpler but would also swallow any
 *  TRAILING flags (`replace … -old -new --all`) into positionals; placeholders don't. */
function normalizeDashLeading(
	args: string[],
	options: ParseArgsOptions,
): { args: string[]; placeholders: Map<string, string> } {
	const entries = Object.entries(options ?? {});
	const stringLong = new Set(
		entries.filter(([, spec]) => spec?.type === "string").map(([name]) => name),
	);
	const stringShort = new Set(
		entries
			.filter(([, spec]) => spec?.type === "string" && spec?.short)
			.map(([, spec]) => spec?.short as string),
	);
	const flagShaped = (token: string): boolean => /^-+[a-zA-Z]/.test(token);
	const consumesNext = (token: string): boolean => {
		if (token.startsWith("--") && !token.includes("="))
			return stringLong.has(token.slice(2));
		if (/^-[a-zA-Z]$/.test(token)) return stringShort.has(token[1] ?? "");
		return false;
	};

	// A NUL byte can never appear in an argv token (the OS null-terminates them),
	// so wrapping the sentinel in NUL makes it impossible for a real argument to
	// collide with a placeholder. Built at runtime to keep NUL out of the source.
	const nul = String.fromCharCode(0);
	const out: string[] = [];
	const placeholders = new Map<string, string>();
	let afterDoubleDash = false;
	let placeholderCount = 0;
	for (let index = 0; index < args.length; index++) {
		const arg = args[index];
		if (arg === undefined) continue;
		if (afterDoubleDash) {
			out.push(arg);
			continue;
		}
		if (arg === "--") {
			afterDoubleDash = true;
			out.push(arg);
			continue;
		}
		if (consumesNext(arg)) {
			const next = args[index + 1];
			// Long `--flag -$value` → `--flag=-$value` so parseArgs keeps the value.
			if (
				next !== undefined &&
				arg.startsWith("--") &&
				next.startsWith("-") &&
				!flagShaped(next)
			) {
				out.push(`${arg}=${next}`);
				index++;
				continue;
			}
			// Otherwise push the flag and its value verbatim — the value is consumed
			// here so it can't be re-read as a dash-led positional below.
			out.push(arg);
			if (next !== undefined) {
				out.push(next);
				index++;
			}
			continue;
		}
		if (arg.startsWith("-") && arg !== "-" && !flagShaped(arg)) {
			// A dash-led positional (`-$500`, `-5`, `-00.00`): shield it behind a
			// NUL-delimited sentinel so parseArgs keeps it a positional, then restore
			// the real text after the parse (see `nul` above — can't collide).
			const placeholder = `${nul}dashpos${placeholderCount++}${nul}`;
			placeholders.set(placeholder, arg);
			out.push(placeholder);
			continue;
		}
		out.push(arg);
	}
	return { args: out, placeholders };
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
			return await fail("BLOCK_NOT_FOUND", err.message, STALE_LOCATOR_HINT);
		}
		throw err;
	}
}
