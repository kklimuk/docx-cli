import {
	CLEARABLE_ATTRS,
	isSectionType,
	type PageGeometry,
	type Run,
	type RunFormat,
	resolveClearTags,
	type SectionType,
} from "@core";
import type { ParagraphOptions } from "@core/blocks";
import {
	firstInvalidRunFormat,
	type RunFormatEnums,
} from "@core/run-formatting";
import { EntryError, fail } from "./respond";

/** Detect the signature of a currency amount whose leading digits were eaten by
 *  the shell. A weak agent double-quotes a `$`-bearing value in bash, and the
 *  shell expands the `$NN` sequence (a positional-param reference) to nothing
 *  BEFORE docx ever sees it: `"$300.00"` → `.00`, `"$10,000"` → `,000`. docx then
 *  faithfully writes the gutted value — the root cause of BOTH "major" currency
 *  bugs in the adversarial review (invoice `.00` cells, contract `,000` redline).
 *  We can't fix bash, but we can refuse the corrupted value at the door.
 *
 *  Returns the offending fragment, or null. The signature is a number fragment
 *  with NO integer part — `[.,]` followed by 2+ digits, not preceded by a digit
 *  (a real `300.00`/`10,000` keeps its integer part) or `$` (legit `$.99` cents).
 *  That makes it self-discriminating: a correctly-passed `$300.00` never trips it,
 *  only the shell-gutted `.00`/`,000` does. */
export function detectShellMangledCurrency(text: string): string | null {
	const match = text.match(/(?<![\d$])([.,]\d{2,})/);
	return match?.[1] ?? null;
}

/** Reject an inline `--text`/`--cells` value that looks shell-gutted, pointing at
 *  single-quoting and `--batch` (both bypass the shell). Returns a fail() exit code
 *  to short-circuit, or null to proceed. Inline argv ONLY — `--batch` values come
 *  from a file, never through the shell, so they're trusted and skip this guard
 *  (which also makes `--batch` the clean escape hatch for a genuine bare `.00`). */
export async function rejectShellMangledValue(
	text: string,
	label = "this value",
): Promise<number | null> {
	const fragment = detectShellMangledCurrency(text);
	if (!fragment) return null;
	return await fail(
		"USAGE",
		`${label} contains "${fragment}" — a number with no integer part, the signature of a "$" amount gutted by the shell (bash turns double-quoted "$300.00" into ".00" and "$10,000" into ",000"). docx would write the corrupted value verbatim.`,
		`Wrap any "$"-bearing value in SINGLE quotes ('$300.00') so bash leaves it alone, or supply it via --batch FILE (JSONL never touches the shell). If you really mean "${fragment}", write its integer part (0${fragment}) or use --batch.`,
	);
}

/** Decode the whitespace escape sequences weak agents type LITERALLY into an
 *  inline `--text` / `--markdown` argv value. A model reaching for a line break
 *  writes `--text "a\nb"`, but bash double-quotes don't interpret `\n` — the CLI
 *  receives the four characters `a`, `\`, `n`, `b`, and the backslash-n lands
 *  verbatim in the run (Word then swallows or literalizes it). We map `\n` / `\r` /
 *  `\r\n` → newline and `\t` → tab so the intended break flows through the same
 *  real-character handling everything else uses (`textToRuns` → `<w:br/>`/`<w:tab/>`,
 *  the markdown parser's block splitting). ONLY these: a `\` before anything else —
 *  markdown's own escaping (`\*`, `\[`, `\.`) and a literal `\\` — passes through
 *  UNTOUCHED, because `\` before a letter was never a valid markdown escape (so
 *  decoding `\n`/`\t`/`\r` can't corrupt markdown) while `\`-before-punctuation is
 *  (so we must not touch it). Applied at the INLINE argv ingress ONLY: `--text-file`
 *  / `--markdown-file` / `--from` read real characters from a file, and `--batch`
 *  JSONL is already decoded by `JSON.parse` — none are re-decoded. The
 *  literal-fidelity channel where a bare `\n` really means backslash-n stays
 *  `--text-file`. */
export function decodeInlineEscapes(value: string): string;
export function decodeInlineEscapes(
	value: string | undefined,
): string | undefined;
export function decodeInlineEscapes(
	value: string | undefined,
): string | undefined {
	if (value === undefined) return undefined;
	return value.replace(/\\r\\n|\\[nrt]/g, (match) =>
		match === "\\t" ? "\t" : "\n",
	);
}

/** Contextual `--help`: a weak agent that reaches for `--text` or `--runs` and
 *  then `--help` gets the slice relevant to what they typed rather than the full
 *  default screen. Runs BEFORE `tryParseArgs` (which short-circuits `--help`
 *  centrally), so it scans raw argv — `--text`/`--runs` are declared options and
 *  `--json` is only a routing hint (undeclared; steers to the runs variant). */
export function pickContextualHelp(
	args: string[],
	variants: { default: string; text: string; runs: string },
): string {
	if (args.includes("--text")) return variants.text;
	if (args.includes("--runs") || args.includes("--json")) return variants.runs;
	return variants.default;
}

/** The tracked-change view selected by the `--accepted` / `--baseline` /
 *  `--current` flags. Shared union across find/replace/read/wc/comments. */
export type View = "accepted" | "baseline" | "current";

/** Resolve the (mutually exclusive) view flags to a single view. Returns `null`
 *  when more than one is set, so the caller can emit its own USAGE error with
 *  the flag wording that command documents. Commands that only expose
 *  `--current`/`--baseline` simply leave `accepted` undefined. */
export function resolveView(values: {
	accepted?: unknown;
	baseline?: unknown;
	current?: unknown;
}): View | null {
	const set =
		(values.accepted ? 1 : 0) +
		(values.baseline ? 1 : 0) +
		(values.current ? 1 : 0);
	if (set > 1) return null;
	if (values.current) return "current";
	if (values.baseline) return "baseline";
	return "accepted";
}

/** Parse a `--runs JSON` argument into a `Run[]`. Shared by insert + edit.
 *  Returns a fail() exit code on malformed JSON or non-array shapes. */
export async function parseRunsArg(json: string): Promise<Run[] | number> {
	let parsed: unknown;
	try {
		parsed = parseJsonLenient(json);
	} catch (jsonError) {
		const message =
			jsonError instanceof Error ? jsonError.message : String(jsonError);
		return fail("USAGE", `Invalid --runs JSON: ${message}`);
	}
	if (!Array.isArray(parsed)) {
		return fail("USAGE", "--runs must be a JSON array of Run objects");
	}
	const runs = parsed as Run[];
	// Enum-valued run formatting (highlight/underline/vertAlign) must be valid
	// or Word writes schema-invalid XML and silently drops it. The markdown
	// `[text]{attrs}` path validates the same way; this closes the gap for the
	// raw `--runs` ingress (shared sets in `@core/run-formatting`).
	for (const run of runs) {
		if (
			run !== null &&
			typeof run === "object" &&
			(run as { type?: unknown }).type === "text"
		) {
			const invalid = firstInvalidRunFormat(run as RunFormatEnums);
			if (invalid) {
				return fail(
					"USAGE",
					`Invalid ${invalid.field} "${invalid.value}" in a --runs text run`,
					`Use ${invalid.valid}.`,
				);
			}
		}
	}
	return runs;
}

type RawValues = Record<
	string,
	string | boolean | (string | boolean)[] | undefined
>;

/** Parse the section flags from a parseArgs result — columns/type PLUS page
 *  geometry (`--orientation`/`--size`/`--margins`). Shared by `docx sections`
 *  and (the geometry slice) `docx create`. Returns a fail() exit
 *  code on invalid values. Page geometry is normalized for `applyPageGeometry`:
 *  `pageSize` is portrait (width ≤ height), and a `WxH` size with `W > H` implies
 *  `orientation: "landscape"` unless `--orientation` says otherwise. */
export async function parseSectionFlags(
	values: RawValues,
): Promise<
	({ columns?: number; sectionType?: SectionType } & PageGeometry) | number
> {
	const out: { columns?: number; sectionType?: SectionType } & PageGeometry =
		{};

	const columnsRaw = values.columns as string | undefined;
	if (columnsRaw !== undefined) {
		// Require pure digits — a bare `Number.parseInt` would silently TRUNCATE
		// "2.5" → 2 and "1e2" → 1, both of which then pass the `> 0` check and write
		// a wrong column count with no error.
		const columns = /^\d+$/.test(columnsRaw.trim())
			? Number.parseInt(columnsRaw, 10)
			: Number.NaN;
		if (!Number.isFinite(columns) || columns <= 0) {
			return fail(
				"USAGE",
				`--columns must be a positive integer, got "${columnsRaw}"`,
			);
		}
		out.columns = columns;
	}

	const sectionTypeRaw = values.type as string | undefined;
	if (sectionTypeRaw !== undefined) {
		if (!isSectionType(sectionTypeRaw)) {
			return fail(
				"USAGE",
				`Invalid --type: ${sectionTypeRaw}`,
				"Valid values: continuous, nextPage, evenPage, oddPage, nextColumn",
			);
		}
		out.sectionType = sectionTypeRaw;
	}

	const orientationRaw = values.orientation as string | undefined;
	if (orientationRaw !== undefined) {
		if (orientationRaw !== "portrait" && orientationRaw !== "landscape") {
			return fail(
				"USAGE",
				`Invalid --orientation: ${orientationRaw}`,
				"Valid values: portrait, landscape",
			);
		}
		out.orientation = orientationRaw;
	}

	const sizeRaw = values.size as string | undefined;
	if (sizeRaw !== undefined) {
		const parsed = parsePageSize(sizeRaw);
		if (isError(parsed)) return fail("USAGE", parsed.error, parsed.hint);
		out.pageSize = { width: parsed.width, height: parsed.height };
		// A WxH size whose width exceeds its height means landscape — honor it
		// unless --orientation was given explicitly.
		if (parsed.impliedLandscape && out.orientation === undefined) {
			out.orientation = "landscape";
		}
	}

	const marginsRaw = values.margins as string | undefined;
	if (marginsRaw !== undefined) {
		const parsed = parseMargins(marginsRaw);
		if (isError(parsed)) return fail("USAGE", parsed.error, parsed.hint);
		out.margins = parsed;
	}

	return out;
}

/** Named page sizes → portrait `{ width, height }` in twips. */
const PAGE_SIZES: Record<string, { width: number; height: number }> = {
	letter: { width: 12240, height: 15840 },
	legal: { width: 12240, height: 20160 },
	tabloid: { width: 15840, height: 24480 },
	ledger: { width: 15840, height: 24480 },
	a3: { width: 16838, height: 23811 },
	a4: { width: 11906, height: 16838 },
	a5: { width: 8391, height: 11906 },
};

/** Parse `--size` — a named size (letter/legal/tabloid/a4/a3/a5) or `WxH` inches
 *  (`8.5x11`, `8.5x11in`). Returns portrait-normalized twips plus whether the
 *  literal `WxH` was landscape-shaped (width > height). */
function parsePageSize(
	raw: string,
):
	| { width: number; height: number; impliedLandscape: boolean }
	| SpacingIndentError {
	const value = raw.trim().toLowerCase();
	const named = PAGE_SIZES[value];
	if (named) {
		return {
			width: named.width,
			height: named.height,
			impliedLandscape: false,
		};
	}
	const match = value.match(
		/^(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*(?:in)?$/,
	);
	if (match) {
		const a = Math.round(
			Number.parseFloat(match[1] as string) * TWIPS_PER_INCH,
		);
		const b = Math.round(
			Number.parseFloat(match[2] as string) * TWIPS_PER_INCH,
		);
		return {
			width: Math.min(a, b),
			height: Math.max(a, b),
			impliedLandscape: a > b,
		};
	}
	return {
		error: `Invalid --size: ${raw}`,
		hint: "Use a name (letter, legal, tabloid, a4, a3, a5) or WxH inches (e.g. 8.5x11in).",
	};
}

/** Parse `--margins` — one inch value (uniform) or four comma-separated values in
 *  CSS order (top,right,bottom,left), matching the `docx:page margins=…` read
 *  note. Inches → twips; signed (a negative margin pulls content into the margin,
 *  which Word allows). A trailing `in` on any part is accepted. */
function parseMargins(
	raw: string,
):
	| { top: number; right: number; bottom: number; left: number }
	| SpacingIndentError {
	const parts = raw
		.trim()
		.split(",")
		.map((part) => part.trim());
	if (parts.length !== 1 && parts.length !== 4) {
		return {
			error: `Invalid --margins: ${raw}`,
			hint: "Use one inch value (uniform, e.g. 1in) or four comma-separated top,right,bottom,left (e.g. 1,1,1,1.5).",
		};
	}
	const twips: number[] = [];
	for (const part of parts) {
		const value = inchesToTwips(part, "margins");
		if (isError(value)) return value;
		twips.push(value);
	}
	if (twips.length === 1) {
		const [uniform] = twips as [number];
		return { top: uniform, right: uniform, bottom: uniform, left: uniform };
	}
	const [top, right, bottom, left] = twips as [number, number, number, number];
	return { top, right, bottom, left };
}

/** Dedupe a list of `--at` comment-id strings while preserving the caller's
 *  order. Normalizes the `cN` prefix (so `--at 3` and `--at c3` collapse to the
 *  same key). Used by the comments batch verbs. */
export function normalizeAndDedupCommentIds(rawIds: string[]): string[] {
	const seen = new Set<string>();
	const ordered: string[] = [];
	for (const raw of rawIds) {
		const normalized = raw.startsWith("c") ? raw : `c${raw}`;
		if (seen.has(normalized)) continue;
		seen.add(normalized);
		ordered.push(normalized);
	}
	return ordered;
}

/** The shared two-line "--batch is the preferred path" help-example intro. The
 *  locator-addressed batch verbs (`edit`, `insert`, `delete`, `comments add`)
 *  state it identically so a weak agent recognizes the same pattern on every
 *  command; only the verb phrase varies. (`replace --batch` words its own —
 *  it's a sed-script whose pitch isn't locator stability.) */
export function batchExampleIntro(what: string): string {
	return `  # ${what} in ONE call — the preferred path (ids never shift
  # mid-batch). Write one JSON object per line to a file, then apply it:`;
}

/** Read a JSONL file (or stdin via `-`) into one parsed object per non-empty
 *  line. Each line must be a JSON object (not an array/scalar); empty lines are
 *  skipped, malformed lines throw with line context so the caller can surface
 *  the failure via `fail("USAGE", ...)`. Shared by every `--batch` ingress
 *  (`comments add/delete/resolve`, `edit`, `insert`, `replace`). */
export async function readJsonlObjects(
	source: string,
): Promise<Record<string, unknown>[]> {
	const raw = await readBatchSource(source);
	const objects: Record<string, unknown>[] = [];
	const lines = raw.split("\n");
	for (let index = 0; index < lines.length; index++) {
		const lineRaw = lines[index];
		if (lineRaw === undefined) continue;
		const line = lineRaw.trim();
		if (line.length === 0) continue;
		let parsed: unknown;
		try {
			parsed = parseJsonLenient(line);
		} catch (error) {
			const message = error instanceof Error ? error.message : String(error);
			throw new Error(`line ${index + 1}: invalid JSON (${message})`);
		}
		if (
			typeof parsed !== "object" ||
			parsed === null ||
			Array.isArray(parsed)
		) {
			throw new Error(`line ${index + 1}: expected a JSON object`);
		}
		objects.push(parsed as Record<string, unknown>);
	}
	return objects;
}

/** `JSON.parse` with one weak-agent recovery. Agents heredoc real TAB (and other
 *  control) characters into string values; JSON forbids raw control chars in
 *  strings, so the text hard-fails — and the observed agent recovery is
 *  substituting spaces, which silently drops a `<w:tab/>` from the document. A
 *  raw control char inside a string can only mean its escaped form, so escape
 *  and retry before rethrowing the ORIGINAL parse error. The shared mechanism
 *  behind every agent-typed JSON ingress: `readJsonlObjects` (`--batch` lines)
 *  and `parseRunsArg` (`--runs`). NOTE the batch path splits its source on `\n`
 *  BEFORE calling this, so a raw NEWLINE inside a string only recovers on the
 *  single-argv `--runs` path — on `--batch` it has already broken into two
 *  malformed lines; every other control char (TAB included) recovers on both. */
function parseJsonLenient(text: string): unknown {
	try {
		return JSON.parse(text);
	} catch (error) {
		try {
			return JSON.parse(escapeControlCharsInStrings(text));
		} catch {
			throw error;
		}
	}
}

/** Escape raw control characters that appear INSIDE JSON string literals
 *  (a literal TAB in a heredoc'd value is the common case) so the line parses.
 *  Characters outside strings are left untouched — structural whitespace is
 *  already legal JSON, and anything else should keep failing loudly. */
function escapeControlCharsInStrings(line: string): string {
	let out = "";
	let inString = false;
	let escaped = false;
	for (const char of line) {
		if (!inString) {
			if (char === '"') inString = true;
			out += char;
			continue;
		}
		if (escaped) {
			out += char;
			escaped = false;
			continue;
		}
		if (char === "\\") {
			out += char;
			escaped = true;
			continue;
		}
		if (char === '"') {
			inString = false;
			out += char;
			continue;
		}
		const code = char.charCodeAt(0);
		if (code < 0x20) {
			// One generic escape covers every control char: JSON.parse reads
			// u-escapes and short escapes identically, and the output is only
			// ever re-parsed, never shown.
			out += `\\u${code.toString(16).padStart(4, "0")}`;
			continue;
		}
		out += char;
	}
	return out;
}

/** Resolve a `--batch` source to its raw JSONL text. Three forms, so a weak agent
 *  can't trip on the on-ramp: `-` reads stdin; a value that LOOKS like inline JSONL
 *  (starts with `{`/`[` after trimming, or contains a newline — a real path does
 *  neither) is used verbatim; otherwise it's a file path. A path that can't be read
 *  fails with a plain message, not the raw ENOENT/ENAMETOOLONG errno an inline value
 *  used to hit (the exact trap that made agents abandon `--batch`). */
async function readBatchSource(source: string): Promise<string> {
	if (source === "-") return await new Response(Bun.stdin.stream()).text();
	if (looksLikeInlineJsonl(source)) return source;
	try {
		return await Bun.file(source).text();
	} catch {
		throw new Error(
			`could not read batch file: ${source} (pass a file path, inline JSONL, or "-" for stdin)`,
		);
	}
}

function looksLikeInlineJsonl(source: string): boolean {
	if (source.includes("\n")) return true;
	const trimmed = source.trimStart();
	return trimmed.startsWith("{") || trimmed.startsWith("[");
}

/** Read a JSONL file (or stdin via `-`) and return the `id` field of each
 *  entry. Each line must be `{"id": "cN"}`; malformed entries throw with line
 *  context. Shared by `comments delete --batch` and `comments resolve --batch`. */
export async function readJsonlIds(source: string): Promise<string[]> {
	const objects = await readJsonlObjects(source);
	const ids: string[] = [];
	for (let index = 0; index < objects.length; index++) {
		const entry = objects[index] as { id?: unknown };
		if (typeof entry.id !== "string" || entry.id.length === 0) {
			throw new Error(`entry ${index}: missing "id"`);
		}
		ids.push(entry.id);
	}
	return ids;
}

/** Normalize a hex color for `<w:rPr>` (`w:color`/`w:shd@w:fill`): strip a single
 *  leading `#` so `--color "#FF0000"` becomes the schema-valid `FF0000` (ST_HexColor
 *  has no `#`). Other values pass through unchanged (Word degrades unknown colors
 *  gracefully). Shared by `edit`, `edit --batch`, and the `styles` verbs. */
export function normalizeHexColor(value: string): string {
	return value.startsWith("#") ? value.slice(1) : value;
}

/** The run-formatting flags (`--bold`/`--color`/`--font`/…), shared by `edit` and
 *  `styles set`/`create`. */
const RUN_FORMAT_FLAGS = [
	"bold",
	"italic",
	"underline",
	"strike",
	"caps",
	"smallcaps",
	"superscript",
	"subscript",
	"color",
	"font",
	"size",
	"highlight",
	"shade",
] as const;

/** Input key → `RunFormat` boolean field. Both the CLI and the batch parser walk
 *  this table, so a new toggle is one row here instead of a hand-written block in
 *  each — the drift these two parsers exist to prevent. */
const BOOLEAN_FORMAT_FIELDS: ReadonlyArray<
	readonly [
		key: string,
		field: "bold" | "italic" | "strike" | "allCaps" | "smallCaps",
	]
> = [
	["bold", "bold"],
	["italic", "italic"],
	["strike", "strike"],
	["caps", "allCaps"],
	["smallcaps", "smallCaps"],
];

/** Input key → `RunFormat` string field plus the normalizer its value goes
 *  through (hex fields drop a leading `#` and upper-case; the rest pass through). */
const STRING_FORMAT_FIELDS: ReadonlyArray<
	readonly [
		key: string,
		field: "color" | "font" | "highlight" | "shade",
		normalize: (value: string) => string,
	]
> = [
	["color", "color", normalizeHexColor],
	["font", "font", (value) => value],
	["highlight", "highlight", (value) => value],
	["shade", "shade", normalizeHexColor],
];

/** True when the invocation carries any run-formatting flag. Booleans default to
 *  `false` from parseArgs (not undefined), so we test against both. */
export function hasRunFormatFlags(values: Record<string, unknown>): boolean {
	return RUN_FORMAT_FLAGS.some(
		(flag) => values[flag] !== undefined && values[flag] !== false,
	);
}

/** Build a `RunFormat` from the formatting flags, or `null` if none are set.
 *  Returns a `{ error }` shape on a bad value (size, mutually-exclusive
 *  super/subscript, or an out-of-range highlight/underline/vertAlign) so the
 *  caller can `fail()`. Shared by `edit` and the `styles` verbs — one parser, so a
 *  style's `--bold`/`--color`/`--size` behaves exactly like a body edit's. */
export function parseRunFormat(
	values: Record<string, unknown>,
): RunFormat | null | { error: string; hint?: string } {
	const format: RunFormat = {};
	for (const [key, field] of BOOLEAN_FORMAT_FIELDS) {
		if (values[key]) format[field] = true;
	}
	for (const [key, field, normalize] of STRING_FORMAT_FIELDS) {
		const value = values[key];
		if (value !== undefined) format[field] = normalize(value as string);
	}
	if (values.underline) format.underline = "single";

	if (values.superscript && values.subscript) {
		return { error: "--superscript and --subscript are mutually exclusive" };
	}
	if (values.superscript) format.vertAlign = "superscript";
	if (values.subscript) format.vertAlign = "subscript";

	if (values.size !== undefined) {
		const points = Number.parseFloat(values.size as string);
		if (!Number.isFinite(points) || points <= 0) {
			return {
				error: `Invalid --size: ${values.size}`,
				hint: "Pass a positive point size, e.g. --size 12 or --size 11.5.",
			};
		}
		format.sizeHalfPoints = Math.round(points * 2);
	}

	const invalid = firstInvalidRunFormat(format);
	if (invalid) {
		return {
			error: `Invalid --${invalid.field}: ${invalid.value}`,
			hint: `Use ${invalid.valid}.`,
		};
	}

	if (Object.keys(format).length === 0) return null;
	return format;
}

/** Reject a single-shot CLI flag passed alongside `--batch`. Every batch surface
 *  reads its per-entry fields from the JSONL, so a CLI flag that duplicates one
 *  is always a mistake — and one sentence teaches the rule everywhere instead of
 *  a slightly different one per command. The `!== false` guard keeps a boolean
 *  flag that merely PARSED as `false` from counting as passed. Returns a fail()
 *  exit code, or undefined when nothing conflicts. */
export async function rejectBatchOnlyFlags(
	values: Record<string, unknown>,
	flags: readonly string[],
	what: string,
	hint: string,
): Promise<number | undefined> {
	const conflicting = flags.find(
		(flag) => values[flag] !== undefined && values[flag] !== false,
	);
	if (conflicting === undefined) return undefined;
	return fail(
		"USAGE",
		`--batch reads each ${what} from the JSONL file; don't also pass --${conflicting} on the CLI`,
		hint,
	);
}

export type BatchValueError = { error: string; hint?: string };

/** Throwing form of `parseBatchBoolean`, for the entry validators that run
 *  inside one try/catch and translate `EntryError` at the top. */
export function requireBatchBoolean(
	raw: Record<string, unknown>,
	index: number,
	field: string,
): boolean {
	const parsed = parseBatchBoolean(raw, index, field);
	if (typeof parsed !== "boolean") {
		throw new EntryError("USAGE", parsed.error, parsed.hint);
	}
	return parsed;
}

/** Read one optional JSONL boolean without truthy coercion. Batch input is a
 * machine format: `"false"` must fail, not silently turn a flag on. */
export function parseBatchBoolean(
	raw: Record<string, unknown>,
	index: number,
	field: string,
): boolean | BatchValueError {
	const value = raw[field];
	if (value === undefined) return false;
	if (typeof value !== "boolean") {
		return { error: `entry ${index}: "${field}" must be a boolean` };
	}
	return value;
}

/** Build a `RunFormat` from a JSONL entry. This is the batch-input twin of
 *  `parseRunFormat`: values are type-checked before use, underline accepts either
 *  `true` (single) or an explicit style string, and errors carry the entry index.
 *  Shared by edit and replace so their batch formatting vocabularies can't drift. */
export function parseBatchRunFormat(
	raw: Record<string, unknown>,
	index: number,
): RunFormat | null | BatchValueError {
	const format: RunFormat = {};
	for (const [key, field] of BOOLEAN_FORMAT_FIELDS) {
		const parsed = parseBatchBoolean(raw, index, key);
		if (typeof parsed !== "boolean") return parsed;
		if (parsed) format[field] = true;
	}
	for (const [key, field, normalize] of STRING_FORMAT_FIELDS) {
		const parsed = batchString(raw, index, key);
		if ("error" in parsed) return parsed;
		if (parsed.value !== undefined) format[field] = normalize(parsed.value);
	}
	if (typeof raw.underline === "string") format.underline = raw.underline;
	else {
		const underline = parseBatchBoolean(raw, index, "underline");
		if (typeof underline !== "boolean") return underline;
		if (underline) format.underline = "single";
	}

	const underlineColor = batchString(raw, index, "underlineColor");
	if ("error" in underlineColor) return underlineColor;
	if (underlineColor.value !== undefined) {
		if (format.underline === undefined) {
			return {
				error: `entry ${index}: "underlineColor" requires "underline"`,
			};
		}
		format.underlineColor = normalizeHexColor(underlineColor.value);
	}

	const superscript = parseBatchBoolean(raw, index, "superscript");
	if (typeof superscript !== "boolean") return superscript;
	const subscript = parseBatchBoolean(raw, index, "subscript");
	if (typeof subscript !== "boolean") return subscript;
	if (superscript && subscript) {
		return {
			error: `entry ${index}: "superscript" and "subscript" are mutually exclusive`,
		};
	}
	if (superscript) format.vertAlign = "superscript";
	if (subscript) format.vertAlign = "subscript";
	if (raw.size !== undefined) {
		const points =
			typeof raw.size === "number"
				? raw.size
				: typeof raw.size === "string" && raw.size.trim().length > 0
					? Number(raw.size)
					: Number.NaN;
		if (!Number.isFinite(points) || points <= 0) {
			return {
				error: `entry ${index}: "size" must be a positive point size (e.g. 12 or 11.5)`,
			};
		}
		format.sizeHalfPoints = Math.round(points * 2);
	}

	const invalid = firstInvalidRunFormat(format);
	if (invalid) {
		return {
			error: `entry ${index}: invalid ${invalid.field} "${invalid.value}"`,
			hint: `Use ${invalid.valid}.`,
		};
	}
	if (Object.keys(format).length === 0) return null;
	return format;
}

/** Parse one `--clear` value into rPr tags. Accepts one name, a comma list,
 * or an array (the shape produced by repeated CLI flags and JSONL entries).
 * `label` names the surface in the error text — a JSONL reader passes the KEY
 * (`"clear"`), because telling an agent editing a batch file about a `--clear`
 * FLAG sends it looking for the wrong thing. */
export function parseClearTags(
	value: unknown,
	label = "--clear",
): Set<string> | BatchValueError {
	const raw =
		typeof value === "string"
			? [value]
			: Array.isArray(value) &&
					value.every((entry) => typeof entry === "string")
				? (value as string[])
				: null;
	if (!raw) {
		return {
			error: `${label} must be an attribute name, a comma list, or an array of names (or "all")`,
		};
	}
	const names = raw
		.flatMap((entry) => entry.split(","))
		.map((name) => name.trim().toLowerCase())
		.filter(Boolean);
	if (names.length === 0) {
		return { error: `${label} needs an attribute name, or "all"` };
	}
	const tags = resolveClearTags(names);
	if (!tags) {
		return {
			error: `Unknown ${label} attribute in "${raw.join(",")}"`,
			hint: `Valid: ${CLEARABLE_ATTRS.join(", ")}, all.`,
		};
	}
	return tags;
}

/** Add JSONL entry context to the shared `--clear` parser. */
export function parseBatchClearTags(
	value: unknown,
	index: number,
): Set<string> | BatchValueError {
	const parsed = parseClearTags(value, '"clear"');
	if (parsed instanceof Set) return parsed;
	return { ...parsed, error: `entry ${index}: ${parsed.error}` };
}

/** Read one optional JSONL integer. Coerces exactly like `size` above (a real
 *  number, or a numeric string — never a boolean or null), so every numeric
 *  batch field agrees on what it accepts and a `{"limit":"3"}` an agent copied
 *  from the `--limit` flag doesn't hard-fail on one surface and work on
 *  another. */
export function parseBatchInteger(
	raw: Record<string, unknown>,
	index: number,
	field: string,
	minimum: number,
): number | undefined | BatchValueError {
	const value = raw[field];
	if (value === undefined) return undefined;
	const parsed =
		typeof value === "number"
			? value
			: typeof value === "string" && value.trim().length > 0
				? Number(value)
				: Number.NaN;
	if (!Number.isInteger(parsed) || parsed < minimum) {
		return {
			error: `entry ${index}: "${field}" must be an integer >= ${minimum}`,
		};
	}
	return parsed;
}

function batchString(
	raw: Record<string, unknown>,
	index: number,
	field: string,
): { value?: string } | BatchValueError {
	const value = raw[field];
	if (value === undefined) return {};
	if (typeof value !== "string") {
		return { error: `entry ${index}: "${field}" must be a string` };
	}
	return { value };
}

const TWIPS_PER_POINT = 20;
const TWIPS_PER_INCH = 1440;
/** `<w:spacing w:line>` units for `lineRule="auto"`: 240ths of a line. */
const LINE_UNITS_PER_MULTIPLE = 240;

type SpacingIndentError = { error: string; hint?: string };

/** Parse the paragraph spacing/indentation flags (shared by `insert` and `edit`,
 *  single-shot and batch) into a `ParagraphOptions` slice, or a `{ error }` the
 *  caller turns into a `fail()`. Units follow the existing CLI conventions:
 *  spacing in points (like font size), indents in inches (like tabs/images), line
 *  spacing as a multiple/alias. `--first-line` and `--hanging` are mutually
 *  exclusive (same OOXML slot). Returns an empty object when no flag is set. */
export function parseSpacingIndentFlags(
	values: Record<string, unknown>,
): Pick<ParagraphOptions, "spacing" | "indent"> | SpacingIndentError {
	const out: Pick<ParagraphOptions, "spacing" | "indent"> = {};
	const spacing: NonNullable<ParagraphOptions["spacing"]> = {};
	const indent: NonNullable<ParagraphOptions["indent"]> = {};

	const before = readMeasure(values, "space-before", pointsToTwips);
	if (isError(before)) return before;
	if (before !== undefined) spacing.before = before;

	const after = readMeasure(values, "space-after", pointsToTwips);
	if (isError(after)) return after;
	if (after !== undefined) spacing.after = after;

	const lineRaw = values["line-spacing"];
	if (lineRaw !== undefined) {
		const parsed = parseLineSpacing(String(lineRaw));
		if (isError(parsed)) return parsed;
		spacing.line = parsed.line;
		spacing.lineRule = parsed.lineRule;
	}

	const left = readMeasure(values, "indent-left", inchesToTwips);
	if (isError(left)) return left;
	if (left !== undefined) indent.left = left;

	const right = readMeasure(values, "indent-right", inchesToTwips);
	if (isError(right)) return right;
	if (right !== undefined) indent.right = right;

	if (values["first-line"] !== undefined && values.hanging !== undefined) {
		return {
			error: "--first-line and --hanging are mutually exclusive",
			hint: "They occupy the same indent slot (positive vs. negative first-line indent). Pass one.",
		};
	}
	const firstLine = readMeasure(values, "first-line", inchesToTwips);
	if (isError(firstLine)) return firstLine;
	if (firstLine !== undefined) indent.firstLine = firstLine;

	const hanging = readMeasure(values, "hanging", unsignedInchesToTwips);
	if (isError(hanging)) return hanging;
	if (hanging !== undefined) indent.hanging = hanging;

	if (Object.keys(spacing).length > 0) out.spacing = spacing;
	if (Object.keys(indent).length > 0) out.indent = indent;
	return out;
}

function isError(value: unknown): value is SpacingIndentError {
	return typeof value === "object" && value !== null && "error" in value;
}

/** Read one measure flag and convert it via `convert`, or undefined if absent. */
function readMeasure(
	values: Record<string, unknown>,
	flag: string,
	convert: (raw: string, flag: string) => number | SpacingIndentError,
): number | undefined | SpacingIndentError {
	const raw = values[flag];
	if (raw === undefined) return undefined;
	return convert(String(raw), flag);
}

/** Points → twips (×20). Accepts a bare number or an explicit `pt` suffix.
 *  Unsigned: `<w:before>`/`<w:after>` are `ST_TwipsMeasure` (non-negative). */
function pointsToTwips(raw: string, flag: string): number | SpacingIndentError {
	const match = raw.trim().match(/^(\d+(?:\.\d+)?)\s*(?:pt)?$/i);
	if (!match) {
		return {
			error: `Invalid --${flag}: ${raw}`,
			hint: "Use a point value, e.g. --space-after 6 (or 6pt).",
		};
	}
	return Math.round(Number.parseFloat(match[1] as string) * TWIPS_PER_POINT);
}

/** Inches → twips (×1440). Accepts a bare number or an explicit `in` suffix, and
 *  (for the signed indent slots) an optional leading `-`. `w:left`/`w:right`/
 *  `w:firstLine` are `ST_SignedTwipsMeasure` — a negative value is a deliberate
 *  outdent into the page margin, which Word produces and the reader surfaces, so
 *  the read→re-apply loop needs to accept it back. `w:hanging` is unsigned
 *  (`signed: false`), as is everything routed through `pointsToTwips`. */
function inchesToTwips(
	raw: string,
	flag: string,
	signed = true,
): number | SpacingIndentError {
	const pattern = signed
		? /^(-?\d+(?:\.\d+)?)\s*(?:in)?$/i
		: /^(\d+(?:\.\d+)?)\s*(?:in)?$/i;
	const match = raw.trim().match(pattern);
	if (!match) {
		return {
			error: `Invalid --${flag}: ${raw}`,
			hint: "Use an inch value, e.g. --indent-left 0.5 (or 0.5in).",
		};
	}
	return Math.round(Number.parseFloat(match[1] as string) * TWIPS_PER_INCH);
}

/** `--hanging` only — the unsigned inch converter (the hanging indent has no
 *  negative form; a negative first-line indent is `--first-line -N`). */
function unsignedInchesToTwips(
	raw: string,
	flag: string,
): number | SpacingIndentError {
	return inchesToTwips(raw, flag, false);
}

type LineRule = NonNullable<ParagraphOptions["spacing"]>["lineRule"];

/** Line-spacing flag → `{ line, lineRule }`. Three forms, mirroring what `read`
 *  emits back in the `docx:p` note so the value round-trips:
 *   • a multiple/alias (`1`, `1.5`, `single`, `double`) → `lineRule="auto"`,
 *     `line` in 240ths of a line;
 *   • `<n>pt` → `lineRule="exact"`, `line` in twips (a fixed line height);
 *   • `<n>pt atLeast` → `lineRule="atLeast"` (a minimum line height).
 *  Word authors exact/atLeast rules; without the pt forms the read note for them
 *  (`line-spacing="18pt exact"`) couldn't be fed back through `--line-spacing`. */
function parseLineSpacing(
	raw: string,
): { line: number; lineRule: LineRule } | SpacingIndentError {
	const value = raw.trim().toLowerCase();
	const ptMatch = value.match(/^(\d+(?:\.\d+)?)\s*pt(?:\s+(exact|atleast))?$/);
	if (ptMatch) {
		const points = Number.parseFloat(ptMatch[1] as string);
		if (points <= 0) return invalidLineSpacing(raw);
		const lineRule: LineRule = ptMatch[2] === "atleast" ? "atLeast" : "exact";
		return { line: Math.round(points * TWIPS_PER_POINT), lineRule };
	}
	const aliases: Record<string, number> = { single: 1, double: 2 };
	const multiple =
		aliases[value] ??
		(/^\d+(?:\.\d+)?$/.test(value) ? Number.parseFloat(value) : Number.NaN);
	if (!Number.isFinite(multiple) || multiple <= 0)
		return invalidLineSpacing(raw);
	return {
		line: Math.round(multiple * LINE_UNITS_PER_MULTIPLE),
		lineRule: "auto",
	};
}

function invalidLineSpacing(raw: string): SpacingIndentError {
	return {
		error: `Invalid --line-spacing: ${raw}`,
		hint: "Use a multiple (1, 1.5, 2), a name (single, double), or an exact point value (e.g. 15pt, or '15pt atLeast').",
	};
}

/** A match's span locator: the classic `p0:3-9` for an in-paragraph match
 *  (`TextMatch` shape), `p0:5-p2:3` when a spanning match crossed boundaries —
 *  the documented cross-paragraph range form (`pN:S-pM:E`), which `comments
 *  add --at` accepts, so a spanning find pipes straight into a spanning
 *  comment. The ONE emitter of span-locator syntax, shared by `find` and
 *  `replace` so every surface prints the same shape. */
export function spanLocator(
	match:
		| { blockId: string; start: number; end: number }
		| {
				startBlockId: string;
				startOffset: number;
				endBlockId: string;
				endOffset: number;
		  },
): string {
	if ("blockId" in match) {
		return `${match.blockId}:${match.start}-${match.end}`;
	}
	if (match.startBlockId === match.endBlockId) {
		return `${match.startBlockId}:${match.startOffset}-${match.endOffset}`;
	}
	return `${match.startBlockId}:${match.startOffset}-${match.endBlockId}:${match.endOffset}`;
}
