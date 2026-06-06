import { isSectionType, type Run, type SectionType } from "@core";
import { fail } from "./respond";

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
	return parsed as Run[];
}

type RawValues = Record<
	string,
	string | boolean | (string | boolean)[] | undefined
>;

/** Parse `--columns N` and `--type T` section flags from a parseArgs result.
 *  Shared by `insert --section` and `edit --at sN`. Returns a fail() exit
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

/** Read a JSONL file (or stdin via `-`) and return the `id` field of each
 *  entry. Each line must be `{"id": "cN"}`; empty lines skipped, malformed
 *  entries throw with line context so the caller can surface the failure
 *  via `fail("USAGE", ...)`. Shared by `comments delete --batch` and
 *  `comments resolve --batch`. */
export async function readJsonlIds(source: string): Promise<string[]> {
	const raw =
		source === "-"
			? await new Response(Bun.stdin.stream()).text()
			: await Bun.file(source).text();
	const ids: string[] = [];
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
		const entry = parsed as { id?: string };
		if (typeof entry.id !== "string" || entry.id.length === 0) {
			throw new Error(`line ${index + 1}: missing "id"`);
		}
		ids.push(entry.id);
	}
	return ids;
}
