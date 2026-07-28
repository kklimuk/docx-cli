import { resolveAuthor, resolveDate, TrackChanges } from "@core";
import {
	type AcrossReplaceResult,
	applyAcrossReplace,
	type FindView,
	findTextSpans,
	type ReplacementFormatting,
	replacementExpander,
	replaceSpanInParagraph,
	selectMatches,
	type TrackedReplaceOptions,
} from "@core/find";
import {
	parseBatchClearTags,
	parseBatchInteger,
	parseBatchRunFormat,
	readJsonlObjects,
	rejectBatchOnlyFlags,
	requireBatchBoolean,
	spanLocator,
} from "../parse-helpers";
import {
	EntryError,
	EXIT,
	fail,
	openOrFail,
	resolveTracked,
	respond,
	respondAck,
} from "../respond";
import { ACROSS_TRACKED_HINT, ACROSS_TRACKED_MESSAGE } from "./across";
import { matchesInScope, ScopeError, validateScopeShape } from "./scope";

type RawValues = Record<
	string,
	string | boolean | (string | boolean)[] | undefined
>;

const BATCH_ENTRY_ONLY_FLAGS = [
	"at",
	"regex",
	"ignore-case",
	"all",
	"limit",
	"current",
	"baseline",
	"exact",
	"clear",
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

/** `docx replace --batch FILE.jsonl`: a sed-script over one read. Each JSONL
 *  line is `{ pattern, replacement, regex?, ignoreCase?, all?, limit?, exact?,
 *  current?|baseline?, author?, clear?, bold?, color?, … }`. Entries apply in
 *  listed order, each
 *  re-finding against the document AS LEFT BY THE PREVIOUS ENTRY (we re-read
 *  the live tree between entries), so later patterns see earlier substitutions
 *  — the same semantics as running `replace` repeatedly. `--dry-run` runs the
 *  whole script in memory and reports per-entry counts without writing. */
export async function runReplaceBatch(
	filePath: string,
	batchSource: string,
	values: RawValues,
): Promise<number> {
	const conflict = await rejectBatchOnlyFlags(
		values,
		BATCH_ENTRY_ONLY_FLAGS,
		"replacement",
		"Put per-entry fields (pattern, replacement, all, formatting, …) on each JSONL line.",
	);
	if (conflict !== undefined) return conflict;

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
		// A "\n" in an entry's pattern/replacement routes it to the cross-paragraph
		// path (JSONL's \n escapes are already real newlines here). Same limits as
		// the single-shot form: no per-paragraph `at` scope, and untracked only —
		// a tracked batch fails the entry rather than record a merge/split wrong.
		if (spec.pattern.includes("\n") || spec.replacement.includes("\n")) {
			if (spec.at !== undefined) {
				return fail(
					"USAGE",
					`entry ${index}: "at" can't scope a multi-line pattern — it spans paragraph boundaries`,
				);
			}
			if (allocator) {
				return fail(
					"USAGE",
					`entry ${index}: ${ACROSS_TRACKED_MESSAGE}`,
					ACROSS_TRACKED_HINT,
				);
			}
			// Applied even under --dry-run: the batch dry-run runs the whole script
			// in memory (later entries must see this one's edits) and only the SAVE
			// is gated.
			let acrossResult: AcrossReplaceResult;
			try {
				acrossResult = applyAcrossReplace(document.body, spec);
			} catch (matcherError) {
				const message =
					matcherError instanceof Error
						? matcherError.message
						: String(matcherError);
				return fail("USAGE", `entry ${index}: invalid pattern: ${message}`);
			}
			results.push({
				pattern: spec.pattern,
				replacement: spec.replacement,
				totalMatches: acrossResult.totalMatches,
				replaced: acrossResult.replaced.length,
				matches: acrossResult.replaced.map((match) => ({
					locator: spanLocator(match),
					text: match.text,
				})),
			});
			if (acrossResult.replaced.length > 0) document.reread();
			continue;
		}
		// Existence check for a scoped entry: validateSpec only checked the `at`
		// SHAPE (no document yet). Resolve it against the LIVE tree here (reread()
		// runs between entries, so ids can shift) and fail loudly on a typo —
		// otherwise a parseable-but-absent locator (p99) would match nothing and
		// the batch would falsely report success, the exact silent no-op the
		// single-shot path errors on (see scope.ts resolveReplaceScope).
		if (spec.at !== undefined) {
			try {
				document.body.resolveBlock(spec.at);
			} catch {
				return fail(
					"BLOCK_NOT_FOUND",
					`entry ${index}: scope "${spec.at}" not found`,
				);
			}
		}
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

		// `at` scopes this entry to one paragraph (the résumé repeated-placeholder
		// fix); applied before first/all/limit so "first" means first in scope.
		const all =
			spec.at !== undefined
				? matchesInScope(findResult.matches, spec.at)
				: findResult.matches;
		const selected = selectMatches(all, spec);

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
		const expand = replacementExpander(spec);
		for (const match of reversed) {
			const blockRef = document.body.resolveBlock(match.blockId);
			replaceSpanInParagraph(
				blockRef.node,
				{ start: match.start, end: match.end },
				expand(match.text),
				tracked,
				spec.view,
				spec.formatting,
			);
		}

		results.push({
			pattern: spec.pattern,
			replacement: spec.replacement,
			totalMatches: all.length,
			replaced: selected.length,
			matches: selected.map((match) => ({
				locator: spanLocator(match),
				text: match.text,
			})),
		});

		// Re-read the live tree so the next entry's find reflects this one's
		// edits (skipped when nothing changed — a zero-match entry leaves the
		// tree untouched, and the full re-walk is the batch loop's biggest cost).
		if (selected.length > 0) document.reread();
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

	// Same no-op principle as single-shot replace: an entry that matched nothing is
	// a silent no-op. Applied entries are already saved (sed-like, sequential), but
	// if ANY entry replaced 0 we exit nonzero so the agent SEES the misses instead
	// of a clean batch ack it reads as total success.
	const noops = results.filter((entry) => entry.replaced === 0);
	if (noops.length > 0) {
		const list = noops.map((entry) => JSON.stringify(entry.pattern)).join(", ");
		const applied = results.length - noops.length;
		// A nonzero batch means "some entries missed" — NOT "nothing changed" (unlike
		// single-shot, which fails before any save). The entries that DID match are
		// already written to disk; spell that out so an agent doesn't revert good edits.
		const savedNote =
			applied === 0
				? "No entry matched, so the document is unchanged."
				: `The other ${applied} ${applied === 1 ? "entry was" : "entries were"} applied and SAVED — this nonzero exit means "some entries missed," not "nothing changed."`;
		return await fail(
			"MATCH_NOT_FOUND",
			`${noops.length} of ${results.length} replace ${results.length === 1 ? "entry" : "entries"} matched nothing (0 occurrences): ${list}. ${savedNote}`,
			`Those patterns weren't found as LITERAL document text — check for read-view markup (\`<mark>\`/\`<u>\`) or a locator in the pattern. \`docx find\` locates the real text; --ignore-case/--regex/--at scope the search.`,
		);
	}

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
	at?: string;
	formatting?: ReplacementFormatting;
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
	const wantCurrent = requireBatchBoolean(raw, index, "current");
	const wantBaseline = requireBatchBoolean(raw, index, "baseline");
	if (wantCurrent && wantBaseline) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: "current" and "baseline" are mutually exclusive`,
		);
	}
	const limitValue = parseBatchInteger(raw, index, "limit", 1);
	if (limitValue !== undefined && typeof limitValue !== "number") {
		throw new EntryError("USAGE", limitValue.error, limitValue.hint);
	}
	const limit = limitValue;
	let at: string | undefined;
	if (raw.at !== undefined) {
		if (typeof raw.at !== "string") {
			throw new EntryError("USAGE", `entry ${index}: "at" must be a string`);
		}
		try {
			at = validateScopeShape(raw.at);
		} catch (error) {
			if (error instanceof ScopeError) {
				throw new EntryError(error.code, `entry ${index}: ${error.message}`);
			}
			throw error;
		}
	}

	const format = parseBatchRunFormat(raw, index);
	if (format && "error" in format) {
		throw new EntryError("USAGE", format.error, format.hint);
	}
	let clearTags: Set<string> | undefined;
	if (raw.clear !== undefined) {
		const parsed = parseBatchClearTags(raw.clear, index);
		if ("error" in parsed) {
			throw new EntryError("USAGE", parsed.error, parsed.hint);
		}
		clearTags = parsed;
	}
	const formatting: ReplacementFormatting | undefined =
		format || clearTags
			? {
					...(clearTags ? { clearTags } : {}),
					...(format ? { format } : {}),
				}
			: undefined;
	const useRegex = requireBatchBoolean(raw, index, "regex");
	const ignoreCaseFlag = requireBatchBoolean(raw, index, "ignoreCase");
	const ignoreCaseAlias = requireBatchBoolean(raw, index, "ignore-case");
	const ignoreCase = ignoreCaseFlag || ignoreCaseAlias;
	const exact = requireBatchBoolean(raw, index, "exact");
	const all = requireBatchBoolean(raw, index, "all");
	let author: string | undefined;
	if (raw.author !== undefined) {
		if (typeof raw.author !== "string") {
			throw new EntryError(
				"USAGE",
				`entry ${index}: "author" must be a string`,
			);
		}
		author = raw.author;
	}

	return {
		pattern: raw.pattern,
		replacement: raw.replacement,
		regex: useRegex,
		ignoreCase,
		exact,
		all,
		...(limit !== undefined ? { limit } : {}),
		view: wantCurrent ? "current" : wantBaseline ? "baseline" : "accepted",
		...(author !== undefined ? { author } : {}),
		...(at !== undefined ? { at } : {}),
		...(formatting ? { formatting } : {}),
	};
}
