import { resolveAuthor, resolveDate, TrackChanges } from "@core";
import {
	type FindView,
	findTextSpans,
	replaceSpanInParagraph,
	type TrackedReplaceOptions,
} from "@core/find";
import { readJsonlObjects } from "../parse-helpers";
import {
	type ErrorCode,
	EXIT,
	fail,
	openOrFail,
	resolveTracked,
	respond,
	respondAck,
} from "../respond";

type RawValues = Record<
	string,
	string | boolean | (string | boolean)[] | undefined
>;

/** `docx replace --batch FILE.jsonl`: a sed-script over one read. Each JSONL
 *  line is `{ pattern, replacement, regex?, ignoreCase?, all?, limit?, exact?,
 *  current?|baseline?, author? }`. Entries apply in listed order, each
 *  re-finding against the document AS LEFT BY THE PREVIOUS ENTRY (we re-read
 *  the live tree between entries), so later patterns see earlier substitutions
 *  — the same semantics as running `replace` repeatedly. `--dry-run` runs the
 *  whole script in memory and reports per-entry counts without writing. */
export async function runReplaceBatch(
	filePath: string,
	batchSource: string,
	values: RawValues,
): Promise<number> {
	const authorFlag = values.author as string | undefined;
	const trackFlag = Boolean(values.track);
	const outputPath = values.output as string | undefined;
	const dryRun = Boolean(values["dry-run"]);

	let rawEntries: Record<string, unknown>[];
	try {
		rawEntries = await readJsonlObjects(batchSource);
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail("USAGE", `Failed to read batch: ${message}`);
	}
	if (rawEntries.length === 0) return fail("USAGE", "Batch file is empty");

	let specs: ReplaceSpec[];
	try {
		specs = rawEntries.map((raw, index) => validateSpec(raw, index));
	} catch (error) {
		if (error instanceof EntryError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	const allocator = resolveTracked(document, trackFlag)
		? new TrackChanges(document).createAllocator()
		: undefined;

	const results: Array<{
		pattern: string;
		replacement: string;
		totalMatches: number;
		replaced: number;
		matches: Array<{ locator: string; text: string }>;
	}> = [];

	for (let index = 0; index < specs.length; index++) {
		const spec = specs[index];
		if (!spec) continue;
		let findResult: ReturnType<typeof findTextSpans>;
		try {
			findResult = findTextSpans(document.body, spec.pattern, {
				regex: spec.regex,
				ignoreCase: spec.ignoreCase,
				view: spec.view,
				exact: spec.exact,
			});
		} catch (matcherError) {
			const message =
				matcherError instanceof Error
					? matcherError.message
					: String(matcherError);
			return fail("USAGE", `entry ${index}: invalid pattern: ${message}`);
		}

		const all = findResult.matches;
		const selected =
			spec.limit !== undefined
				? all.slice(0, spec.limit)
				: spec.all
					? all
					: all.slice(0, 1);

		const tracked: TrackedReplaceOptions | undefined = allocator
			? {
					meta: {
						author: resolveAuthor(spec.author ?? authorFlag),
						date: resolveDate(),
					},
					allocator,
				}
			: undefined;

		// Reverse document order so earlier offsets stay valid as later ones get
		// rewritten — same as single-shot replace.
		const reversed = [...selected].sort((left, right) => {
			if (left.blockId !== right.blockId) {
				return right.blockId.localeCompare(left.blockId);
			}
			return right.start - left.start;
		});
		const regexFlags = spec.ignoreCase ? "i" : "";
		for (const match of reversed) {
			const concreteReplacement = spec.regex
				? match.text.replace(
						new RegExp(spec.pattern, regexFlags),
						spec.replacement,
					)
				: spec.replacement;
			const blockRef = document.body.resolveBlock(match.blockId);
			replaceSpanInParagraph(
				blockRef.node,
				{ start: match.start, end: match.end },
				concreteReplacement,
				tracked,
				spec.view,
			);
		}

		results.push({
			pattern: spec.pattern,
			replacement: spec.replacement,
			totalMatches: all.length,
			replaced: selected.length,
			matches: selected.map((match) => ({
				locator: `${match.blockId}:${match.start}-${match.end}`,
				text: match.text,
			})),
		});

		// Re-read the live tree so the next entry's find reflects this one's edits.
		document.reread();
	}

	if (dryRun) {
		await respond({
			operation: "replace",
			dryRun: true,
			path: filePath,
			batch: results,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	await document.save(outputPath);

	await respondAck({
		ok: true,
		operation: "replace",
		path: outputPath ?? filePath,
		batch: results,
	});
	return EXIT.OK;
}

type ReplaceSpec = {
	pattern: string;
	replacement: string;
	regex: boolean;
	ignoreCase: boolean;
	exact: boolean;
	all: boolean;
	limit?: number;
	view: FindView;
	author?: string;
};

function validateSpec(
	raw: Record<string, unknown>,
	index: number,
): ReplaceSpec {
	if (typeof raw.pattern !== "string" || raw.pattern.length === 0) {
		throw new EntryError("USAGE", `entry ${index}: "pattern" is required`);
	}
	if (typeof raw.replacement !== "string") {
		throw new EntryError(
			"USAGE",
			`entry ${index}: "replacement" is required (use "" to delete the match)`,
		);
	}
	const wantCurrent = Boolean(raw.current);
	const wantBaseline = Boolean(raw.baseline);
	if (wantCurrent && wantBaseline) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: "current" and "baseline" are mutually exclusive`,
		);
	}
	let limit: number | undefined;
	if (raw.limit !== undefined) {
		const value = Number(raw.limit);
		if (!Number.isInteger(value) || value <= 0) {
			throw new EntryError(
				"USAGE",
				`entry ${index}: "limit" must be a positive integer`,
			);
		}
		limit = value;
	}
	return {
		pattern: raw.pattern,
		replacement: raw.replacement,
		regex: Boolean(raw.regex),
		ignoreCase: Boolean(raw.ignoreCase ?? raw["ignore-case"]),
		exact: Boolean(raw.exact),
		all: Boolean(raw.all),
		...(limit !== undefined ? { limit } : {}),
		view: wantCurrent ? "current" : wantBaseline ? "baseline" : "accepted",
		...(typeof raw.author === "string" ? { author: raw.author } : {}),
	};
}

/** Per-entry validation failure, mirroring `comments add --batch`. */
class EntryError extends Error {
	constructor(
		public code: ErrorCode,
		message: string,
		public hint?: string,
	) {
		super(message);
		this.name = "EntryError";
	}
}
