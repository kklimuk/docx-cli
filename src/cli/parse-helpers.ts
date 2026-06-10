import { isSectionType, type Run, type SectionType } from "@core";
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
