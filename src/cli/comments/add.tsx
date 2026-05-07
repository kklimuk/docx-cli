import {
	type DocView,
	type Locator,
	locatorToBlockTarget,
	parseLocator,
	saveDocView,
} from "@core";
import { type FindView, findTextSpans } from "@core/find";
import { parseArgs } from "util";
import {
	type ErrorCode,
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
	respond,
	respondAck,
	setVerboseAck,
	writeStdout,
} from "../respond";
import {
	addCommentMarkersToParagraph,
	addCommentRangeMarkers,
	authorInitials,
	CommentBody,
	type CommentSpan,
	ensureCommentsPart,
	generateParaId,
	nextCommentId,
	SpanOutOfRangeError,
} from "./helpers";

const HELP = `docx comments add — anchor a new comment to a locator or phrase

Usage:
  docx comments add FILE --range LOCATOR --text TEXT [options]
  docx comments add FILE --anchor PHRASE --text TEXT [options]
  docx comments add FILE --batch FILE.jsonl [options]
  docx comments add FILE --batch -    [options]   # read JSONL from stdin

Anchor (one required, mutually exclusive):
  --range LOCATOR     Where to anchor. Supports:
                        pN              whole paragraph
                        pN:S-E          chars S..E of pN
                        pN:S-pM:E       chars S of pN through char E of pM
                        tT:rRcC:pK      whole cell paragraph
                        tT:rRcC:pK:S-E  chars S..E of cell paragraph
  --anchor PHRASE     Find the phrase via the same matcher as \`docx find\`
                      (default: accepted view, normalization on). The match
                      is converted to a pN:S-E locator and used as the
                      anchor. Errors if the phrase matches more than once
                      without --occurrence.
  --batch PATH        Read a JSONL file (one entry per line). Each entry is
                      a JSON object with the same shape as a single-shot
                      call: { range | anchor (+ optional occurrence), text,
                      optional author }. Use - for stdin.

Per-entry fields:
  --text TEXT         Comment body. Required for single-shot; required as
                      a "text" field on every JSONL entry.
  --occurrence N      For --anchor: pick the Nth match (1-indexed,
                      default 1). Errors if N is out of range.
  --author NAME       Author name (default: $DOCX_AUTHOR)

View (for resolving --range offsets and --anchor matches):
  --current           Resolve offsets in the raw concatenation (with both
                      ins and del text). Use when offsets came from
                      \`docx find --current\` or hand-counted bytes.
  --baseline          Resolve offsets in the pre-change view (skip ins/
                      moveTo).
                      Default: accepted view (skip del/moveFrom) — matches
                      \`find\`'s default. Mutually exclusive.

General options:
  -o, --output PATH   Write to PATH instead of overwriting FILE
  --dry-run           Print what would be added; do not write the file
  -v, --verbose       Print the success ack JSON (default: silent on success
                      for single; batch always prints the minted ids)
  -h, --help          Show this help

Examples:
  docx comments add doc.docx --range p3 --text "Reconsider this paragraph"
  docx comments add doc.docx --anchor "fatally flawed" --text "Cite source?"
  docx comments add doc.docx --anchor "TODO" --occurrence 2 --text "Pick up here"
  docx comments add doc.docx --batch reviews.jsonl
  cat reviews.jsonl | docx comments add doc.docx --batch -

Batch JSONL example:
  {"range": "p3:5-20", "text": "Sharper wording?"}
  {"anchor": "fatally flawed", "text": "Cite Bianco here.", "author": "Reviewer"}
`;

type RawEntry = {
	range?: string;
	anchor?: string;
	/** Explicit `--occurrence N` (or `"occurrence": N` in JSONL). When unset
	 *  AND the anchor matches more than once, we error so the agent knows to
	 *  disambiguate. When set, we use the Nth match (1-indexed). */
	occurrence?: number;
	text: string;
	author?: string;
};

type ResolvedEntry =
	| {
			kind: "single";
			blockId: string;
			span?: CommentSpan;
			text: string;
			author: string;
			locatorString: string;
	  }
	| {
			kind: "range";
			startBlockId: string;
			startOffset: number;
			endBlockId: string;
			endOffset: number;
			text: string;
			author: string;
			locatorString: string;
	  };

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

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				range: { type: "string" },
				anchor: { type: "string" },
				occurrence: { type: "string" },
				text: { type: "string" },
				batch: { type: "string" },
				author: { type: "string" },
				current: { type: "boolean" },
				baseline: { type: "boolean" },
				output: { type: "string", short: "o" },
				"dry-run": { type: "boolean" },
				verbose: { type: "boolean", short: "v" },
				help: { type: "boolean", short: "h" },
			},
		});
	} catch (parseError) {
		const message =
			parseError instanceof Error ? parseError.message : String(parseError);
		return fail("USAGE", message, HELP);
	}

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const rangeInput = parsed.values.range as string | undefined;
	const anchorInput = parsed.values.anchor as string | undefined;
	const batchInput = parsed.values.batch as string | undefined;
	const text = parsed.values.text as string | undefined;
	const occurrenceRaw = parsed.values.occurrence as string | undefined;
	const defaultAuthor =
		(parsed.values.author as string | undefined) ?? Bun.env.DOCX_AUTHOR ?? "";
	const outputPath = parsed.values.output as string | undefined;
	const dryRun = Boolean(parsed.values["dry-run"]);
	const wantCurrent = Boolean(parsed.values.current);
	const wantBaseline = Boolean(parsed.values.baseline);
	if (wantCurrent && wantBaseline) {
		return fail("USAGE", "--current and --baseline are mutually exclusive");
	}
	const findView: FindView = wantCurrent
		? "current"
		: wantBaseline
			? "baseline"
			: "accepted";

	const anchorCount =
		(rangeInput ? 1 : 0) + (anchorInput ? 1 : 0) + (batchInput ? 1 : 0);
	if (anchorCount === 0) {
		return fail(
			"USAGE",
			"Specify exactly one of --range, --anchor, or --batch",
			HELP,
		);
	}
	if (anchorCount > 1) {
		return fail(
			"USAGE",
			"--range, --anchor, and --batch are mutually exclusive",
			HELP,
		);
	}

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	let rawEntries: RawEntry[];
	if (batchInput) {
		if (text !== undefined || rangeInput || anchorInput || occurrenceRaw) {
			return fail(
				"USAGE",
				"--batch reads each entry's range/anchor/text/author from JSONL — do not pass them on the CLI",
				HELP,
			);
		}
		try {
			rawEntries = await readJsonlEntries(batchInput);
		} catch (error) {
			const message = error instanceof Error ? error.message : String(error);
			return fail("USAGE", `Failed to read batch: ${message}`);
		}
		if (rawEntries.length === 0) {
			return fail("USAGE", "Batch file is empty");
		}
	} else {
		if (!text) return fail("USAGE", "Missing --text TEXT", HELP);
		const occurrence =
			occurrenceRaw === undefined ? undefined : Number(occurrenceRaw);
		if (
			occurrence !== undefined &&
			(!Number.isInteger(occurrence) || occurrence < 1)
		) {
			return fail(
				"USAGE",
				`--occurrence must be a positive integer (1-indexed), got "${occurrenceRaw}"`,
			);
		}
		rawEntries = [
			{
				range: rangeInput,
				anchor: anchorInput,
				occurrence,
				text,
			},
		];
	}

	let resolved: ResolvedEntry[];
	try {
		resolved = rawEntries.map((entry, index) =>
			resolveEntry(
				view,
				entry,
				defaultAuthor,
				index,
				batchInput !== undefined,
				findView,
			),
		);
	} catch (error) {
		if (error instanceof EntryError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	if (dryRun) {
		await respond({
			ok: true,
			operation: "comments.add",
			dryRun: true,
			path,
			...(outputPath ? { output: outputPath } : {}),
			batch: resolved.map((entry) => ({
				locator: entry.locatorString,
			})),
		});
		return EXIT.OK;
	}

	const date = new Date().toISOString();
	const minted: Array<{ commentId: string; locator: string }> = [];

	for (const entry of resolved) {
		const numericId = nextCommentId(view);
		const paraId = generateParaId();

		try {
			if (entry.kind === "single") {
				const paragraphRef = await resolveBlockOrFail(view, entry.blockId);
				if (typeof paragraphRef === "number") {
					return paragraphRef;
				}
				addCommentMarkersToParagraph(
					paragraphRef.node,
					numericId,
					entry.span,
					findView,
				);
			} else {
				const startRef = await resolveBlockOrFail(view, entry.startBlockId);
				if (typeof startRef === "number") return startRef;
				const endRef = await resolveBlockOrFail(view, entry.endBlockId);
				if (typeof endRef === "number") return endRef;
				addCommentRangeMarkers(
					startRef.node,
					entry.startOffset,
					endRef.node,
					entry.endOffset,
					numericId,
					findView,
				);
			}
		} catch (error) {
			if (error instanceof SpanOutOfRangeError) {
				return fail("INVALID_LOCATOR", error.message);
			}
			throw error;
		}

		const commentsRoot = ensureCommentsPart(view);
		commentsRoot.children.push(
			<CommentBody
				options={{
					id: numericId,
					author: entry.author,
					date,
					initials: authorInitials(entry.author),
					paraId,
					text: entry.text,
				}}
			/>,
		);

		minted.push({
			commentId: `c${numericId}`,
			locator: entry.locatorString,
		});
	}

	await saveDocView(view, outputPath);

	if (batchInput) {
		// Batch: always print the minted ids — the agent can't reconstruct
		// them without re-reading. Verbose adds the full envelope.
		await respond({
			ok: true,
			operation: "comments.add",
			path: outputPath ?? path,
			batch: minted,
		});
	} else {
		const single = minted[0];
		if (!single) {
			// Unreachable: a non-batch invocation always produces exactly one
			// entry by construction. Defensive narrow so we don't ! the array.
			throw new Error("internal: single-shot path produced no minted entry");
		}
		await respondAck({
			ok: true,
			operation: "comments.add",
			path: outputPath ?? path,
			commentId: single.commentId,
			locator: single.locator,
		});
	}
	return EXIT.OK;
}

function resolveEntry(
	view: DocView,
	raw: RawEntry,
	defaultAuthor: string,
	entryIndex: number,
	isBatch: boolean,
	findView: FindView,
): ResolvedEntry {
	const labelPrefix = isBatch ? `entry ${entryIndex}: ` : "";
	const author = raw.author ?? defaultAuthor;

	const anchorCount = (raw.range ? 1 : 0) + (raw.anchor ? 1 : 0);
	if (anchorCount === 0) {
		throw new EntryError(
			"USAGE",
			`${labelPrefix}must specify either "range" or "anchor"`,
		);
	}
	if (anchorCount > 1) {
		throw new EntryError(
			"USAGE",
			`${labelPrefix}"range" and "anchor" are mutually exclusive`,
		);
	}
	if (typeof raw.text !== "string" || raw.text.length === 0) {
		throw new EntryError("USAGE", `${labelPrefix}"text" is required`);
	}

	if (raw.range) {
		return resolveLocatorEntry(raw.range, raw.text, author, labelPrefix);
	}

	const anchor = raw.anchor;
	if (!anchor) {
		// Unreachable given the anchorCount check above, but keeps the type
		// narrowed without a non-null assertion.
		throw new EntryError("USAGE", `${labelPrefix}missing anchor`);
	}
	const occurrenceExplicit = raw.occurrence !== undefined;
	const occurrence = raw.occurrence ?? 1;
	if (!Number.isInteger(occurrence) || occurrence < 1) {
		throw new EntryError(
			"USAGE",
			`${labelPrefix}"occurrence" must be a positive integer, got ${JSON.stringify(raw.occurrence)}`,
		);
	}
	return resolveAnchorEntry(
		view,
		anchor,
		occurrence,
		occurrenceExplicit,
		raw.text,
		author,
		labelPrefix,
		findView,
	);
}

function resolveLocatorEntry(
	rangeInput: string,
	text: string,
	author: string,
	labelPrefix: string,
): ResolvedEntry {
	let locator: Locator;
	try {
		locator = parseLocator(rangeInput);
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		throw new EntryError("INVALID_LOCATOR", `${labelPrefix}${message}`);
	}

	if (locator.kind === "range") {
		return {
			kind: "range",
			startBlockId: locator.start.blockId,
			startOffset: locator.start.offset,
			endBlockId: locator.end.blockId,
			endOffset: locator.end.offset,
			text,
			author,
			locatorString: rangeInput,
		};
	}

	const target = locatorToBlockTarget(locator);
	if (!target) {
		throw new EntryError(
			"INVALID_LOCATOR",
			`${labelPrefix}comments add supports paragraph locators only (pN, pN:start-end, pN:S-pM:E, or tT:rRcC:pK[:start-end])`,
			"Comments and images are not valid anchors.",
		);
	}
	return {
		kind: "single",
		blockId: target.blockId,
		span: target.span,
		text,
		author,
		locatorString: rangeInput,
	};
}

function resolveAnchorEntry(
	view: DocView,
	anchor: string,
	occurrence: number,
	occurrenceExplicit: boolean,
	text: string,
	author: string,
	labelPrefix: string,
	findView: FindView,
): ResolvedEntry {
	const result = findTextSpans(view.doc, anchor, { view: findView });
	const matches = result.matches;
	if (matches.length === 0) {
		throw new EntryError(
			"MATCH_NOT_FOUND",
			`${labelPrefix}anchor not found: ${JSON.stringify(anchor)}`,
			result.normalizedQuery
				? `Searched as "${result.normalizedQuery}" after normalization (${(result.normalizationApplied ?? []).join(", ")}).`
				: undefined,
		);
	}
	if (matches.length > 1 && !occurrenceExplicit) {
		throw new EntryError(
			"USAGE",
			`${labelPrefix}anchor matches ${matches.length} times; pass --occurrence N (1..${matches.length}) to disambiguate`,
		);
	}
	if (occurrence > matches.length) {
		throw new EntryError(
			"MATCH_NOT_FOUND",
			`${labelPrefix}only ${matches.length} match(es) for anchor; --occurrence ${occurrence} is out of range`,
		);
	}
	const match = matches[occurrence - 1];
	if (!match) {
		throw new EntryError(
			"MATCH_NOT_FOUND",
			`${labelPrefix}match index ${occurrence} not found`,
		);
	}
	const locatorString = `${match.blockId}:${match.start}-${match.end}`;
	return {
		kind: "single",
		blockId: match.blockId,
		span: { start: match.start, end: match.end },
		text,
		author,
		locatorString,
	};
}

async function readJsonlEntries(source: string): Promise<RawEntry[]> {
	const raw =
		source === "-"
			? await new Response(Bun.stdin.stream()).text()
			: await Bun.file(source).text();
	const entries: RawEntry[] = [];
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
		entries.push(parsed as RawEntry);
	}
	return entries;
}
