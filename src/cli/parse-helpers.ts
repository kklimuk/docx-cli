import { isSectionType, type Run, type SectionType } from "@core";
import type { ParagraphOptions } from "@core/blocks";
import {
	firstInvalidRunFormat,
	type RunFormatEnums,
} from "@core/run-formatting";
import { fail } from "./respond";

/** Detect markdown-looking syntax in a value meant for `--text`, which writes
 *  LITERAL text. Weak agents conflate `--text` with `--markdown`, bake literal
 *  `**`/`#`/`[](…)` into runs, then try to scrub them with `replace` — the cascade
 *  that produced the résumé blocker. Returns a short description of the construct
 *  found, or null. HIGH-CONFIDENCE patterns only (paired emphasis, a leading ATX
 *  heading, a link), so literal prose with a stray asterisk doesn't trip it.
 *  Shared by `insert --text` and `edit --text`. */
export function detectMarkdownInPlainText(text: string): string | null {
	if (/\*\*[^*\n]+\*\*/.test(text)) return "bold (**…**)";
	if (/__[^_\n]+__/.test(text)) return "bold (__…__)";
	if (/(^|\n) {0,3}#{1,6}\s/.test(text)) return "a heading (#…)";
	if (/\[[^\]\n]+\]\([^)\n]+\)/.test(text)) return "a link ([text](url))";
	return null;
}

/** Reject a `--text` value that looks like markdown, pointing at the right verb.
 *  Returns a fail() exit code to short-circuit, or null to proceed. */
export async function rejectMarkdownInText(
	text: string,
	help: string,
): Promise<number | null> {
	const found = detectMarkdownInPlainText(text);
	if (!found) return null;
	return await fail(
		"USAGE",
		`--text writes literal characters, but this value looks like markdown: ${found}. It would be baked in verbatim (e.g. literal ** around the word), not rendered.`,
		`Use --markdown to parse it (handles ${found}, headings, lists, links), --bold/--italic for a uniformly-emphasized run, or --runs for literal text that really contains these characters. Help:\n${help}`,
	);
}

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
	help: string,
	label = "this value",
): Promise<number | null> {
	const fragment = detectShellMangledCurrency(text);
	if (!fragment) return null;
	return await fail(
		"USAGE",
		`${label} contains "${fragment}" — a number with no integer part, the signature of a "$" amount gutted by the shell (bash turns double-quoted "$300.00" into ".00" and "$10,000" into ",000"). docx would write the corrupted value verbatim.`,
		`Wrap any "$"-bearing value in SINGLE quotes ('$300.00') so bash leaves it alone, or supply it via --batch FILE (JSONL never touches the shell). If you really mean "${fragment}", write its integer part (0${fragment}) or use --batch. Help:\n${help}`,
	);
}

/** Parse `--task` value into a boolean (checked) or null if unrecognized.
 *  Accepts `checked`/`unchecked` (canonical) plus a few short forms agents
 *  reach for naturally. Shared by `insert --task` and `edit --at pN --task`. */
export function parseTaskFlag(value: string): boolean | null {
	const normalized = value.toLowerCase();
	if (normalized === "checked" || normalized === "true" || normalized === "1")
		return true;
	if (
		normalized === "unchecked" ||
		normalized === "false" ||
		normalized === "0"
	)
		return false;
	return null;
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
		parsed = JSON.parse(json);
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

/** Parse `--columns N` and `--type T` section flags from a parseArgs result.
 *  Shared by `docx sections` and `edit --at sN`. Returns a fail() exit
 *  code on invalid values. */
export async function parseSectionFlags(
	values: RawValues,
): Promise<{ columns?: number; sectionType?: SectionType } | number> {
	const out: { columns?: number; sectionType?: SectionType } = {};

	const columnsRaw = values.columns as string | undefined;
	if (columnsRaw !== undefined) {
		const columns = Number.parseInt(columnsRaw, 10);
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

	return out;
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

/** Read a JSONL file (or stdin via `-`) into one parsed object per non-empty
 *  line. Each line must be a JSON object (not an array/scalar); empty lines are
 *  skipped, malformed lines throw with line context so the caller can surface
 *  the failure via `fail("USAGE", ...)`. Shared by every `--batch` ingress
 *  (`comments add/delete/resolve`, `edit`, `insert`, `replace`). */
export async function readJsonlObjects(
	source: string,
): Promise<Record<string, unknown>[]> {
	const raw =
		source === "-"
			? await new Response(Bun.stdin.stream()).text()
			: await Bun.file(source).text();
	const objects: Record<string, unknown>[] = [];
	const lines = raw.split("\n");
	for (let index = 0; index < lines.length; index++) {
		const lineRaw = lines[index];
		if (lineRaw === undefined) continue;
		const line = lineRaw.trim();
		if (line.length === 0) continue;
		let parsed: unknown;
		try {
			parsed = JSON.parse(line);
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
