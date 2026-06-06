import {
	type CommentAnchorSpec,
	Comments,
	CommentsError,
	type Document,
	describeForms,
	type Locator,
	locatorToBlockTarget,
	parseLocator,
} from "@core";
import { type FindView, findTextSpans } from "@core/find";
import {
	type ErrorCode,
	EXIT,
	fail,
	openOrFail,
	respond,
	respondMinted,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const AT_FORMS = describeForms(
	["paragraph", "span", "crossSpan", "cellParagraph", "cellSpan"],
	"                        ",
);

const HELP = `docx comments add — anchor a new comment to a locator or phrase

Usage:
  docx comments add FILE --at LOCATOR --text TEXT [options]
  docx comments add FILE --anchor PHRASE --text TEXT [options]
  docx comments add FILE --batch FILE.jsonl [options]
  docx comments add FILE --batch -    [options]   # read JSONL from stdin

Anchor (one required, mutually exclusive):
  --at LOCATOR        Where to anchor. Supports:
${AT_FORMS}
  --anchor PHRASE     Find the phrase via the same matcher as \`docx find\`
                      (default: accepted view, normalization on). The match is
                      converted to a pN:S-E locator and used as the anchor.
                      Errors if the phrase matches more than once without
                      --occurrence.
  --batch PATH        Read a JSONL file (one entry per line). Each entry is a
                      JSON object: { at | anchor (+ optional occurrence), text,
                      optional author }. Use - for stdin.

Per-entry fields:
  --text TEXT         Comment body. Required for single-shot; required as a
                      "text" field on every JSONL entry.
  --occurrence N      For --anchor: pick the Nth match (1-indexed, default 1).
                      Errors if N is out of range.
  --author NAME       Author name (default: $DOCX_AUTHOR)

View — for resolving --at offsets and --anchor matches (mutually exclusive;
default: accepted view, matching \`docx find\`):
  --current           Resolve offsets in the raw concatenation (ins+del text).
                      Use when offsets came from \`docx find --current\`.
  --baseline          Resolve offsets in the pre-change document (skip ins/moveTo).

General options:
  -o, --output PATH   Write to PATH instead of overwriting FILE
  --dry-run           Print what would be added; do not write the file
  -v, --verbose       Print the full success ack JSON
  -h, --help          Show this help

Output:
  Prints the new comment id (e.g. c0), one per line for --batch. Address it
  later with \`--at c0\`. --verbose prints the full ack
  {ok:true, operation, path, commentId, locator}. Errors print
  {code, error, hint?} with a nonzero exit. Notation: uppercase letters are
  numeric indices; offsets are 0-based, end-exclusive. See \`docx info locators\`.
  Discover existing comment ids with \`docx comments list FILE\`.

Examples:
  docx comments add doc.docx --at p3 --text "Reconsider this paragraph"
  docx comments add doc.docx --at p3:5-20 --text "Tighten this clause"
  docx comments add doc.docx --anchor "fatally flawed" --text "Cite source?"
  docx comments add doc.docx --anchor "TODO" --occurrence 2 --text "Pick up here"
  docx comments add doc.docx --batch reviews.jsonl
  cat reviews.jsonl | docx comments add doc.docx --batch -

Batch JSONL example:
  {"at": "p3:5-20", "text": "Sharper wording?"}
  {"anchor": "fatally flawed", "text": "Cite Bianco here.", "author": "Reviewer"}
`;

type RawEntry = {
	at?: string;
	anchor?: string;
	/** Explicit `--occurrence N` (or `"occurrence": N` in JSONL). When unset
	 *  AND the anchor matches more than once, we error so the agent knows to
	 *  disambiguate. When set, we use the Nth match (1-indexed). */
	occurrence?: number;
	text: string;
	author?: string;
};

type ResolvedEntry = {
	spec: CommentAnchorSpec;
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
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
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
		HELP,
	);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const atInput = parsed.values.at as string | undefined;
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
		(atInput ? 1 : 0) + (anchorInput ? 1 : 0) + (batchInput ? 1 : 0);
	if (anchorCount === 0) {
		return fail(
			"USAGE",
			"Specify exactly one of --at, --anchor, or --batch",
			HELP,
		);
	}
	if (anchorCount > 1) {
		return fail(
			"USAGE",
			"--at, --anchor, and --batch are mutually exclusive",
			HELP,
		);
	}

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	let rawEntries: RawEntry[];
	if (batchInput) {
		if (text !== undefined || atInput || anchorInput || occurrenceRaw) {
			return fail(
				"USAGE",
				"--batch reads each entry's at/anchor/text/author from JSONL — do not pass them on the CLI",
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
				at: atInput,
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
				document,
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

	const lens = new Comments(document);
	const minted: Array<{ commentId: string; locator: string }> = [];

	for (let index = 0; index < resolved.length; index++) {
		const entry = resolved[index];
		if (!entry) continue;
		try {
			const numericId = lens.add(entry.spec, {
				body: entry.text,
				author: entry.author,
				findView,
			});
			minted.push({
				commentId: `c${numericId}`,
				locator: entry.locatorString,
			});
		} catch (error) {
			if (error instanceof CommentsError) {
				const prefix = batchInput !== undefined ? `entry ${index}: ` : "";
				return fail(error.code, `${prefix}${error.message}`, error.hint);
			}
			throw error;
		}
	}

	await document.save(outputPath);

	// Minted comment ids are the addressable handle (`--at cN`) and can't be
	// reconstructed without re-reading, so they print by default — one `cN`
	// per line — upgrading to the full ack under --verbose.
	if (batchInput) {
		await respondMinted(
			minted.map((entry) => entry.commentId),
			{
				ok: true,
				operation: "comments.add",
				path: outputPath ?? path,
				batch: minted,
			},
		);
	} else {
		const single = minted[0];
		if (!single) {
			// Unreachable: a non-batch invocation always produces exactly one
			// entry by construction. Defensive narrow so we don't ! the array.
			throw new Error("internal: single-shot path produced no minted entry");
		}
		await respondMinted([single.commentId], {
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
	document: Document,
	raw: RawEntry,
	defaultAuthor: string,
	entryIndex: number,
	isBatch: boolean,
	findView: FindView,
): ResolvedEntry {
	const labelPrefix = isBatch ? `entry ${entryIndex}: ` : "";
	const author = raw.author ?? defaultAuthor;

	const anchorCount = (raw.at ? 1 : 0) + (raw.anchor ? 1 : 0);
	if (anchorCount === 0) {
		throw new EntryError(
			"USAGE",
			`${labelPrefix}must specify either "at" or "anchor"`,
		);
	}
	if (anchorCount > 1) {
		throw new EntryError(
			"USAGE",
			`${labelPrefix}"at" and "anchor" are mutually exclusive`,
		);
	}
	if (typeof raw.text !== "string" || raw.text.length === 0) {
		throw new EntryError("USAGE", `${labelPrefix}"text" is required`);
	}

	if (raw.at) {
		return resolveLocatorEntry(raw.at, raw.text, author, labelPrefix);
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
		document,
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
	atValue: string,
	text: string,
	author: string,
	labelPrefix: string,
): ResolvedEntry {
	let locator: Locator;
	try {
		locator = parseLocator(atValue);
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		throw new EntryError("INVALID_LOCATOR", `${labelPrefix}${message}`);
	}

	if (locator.kind === "range") {
		return {
			spec: {
				kind: "range",
				startBlockId: locator.start.blockId,
				startOffset: locator.start.offset,
				endBlockId: locator.end.blockId,
				endOffset: locator.end.offset,
			},
			text,
			author,
			locatorString: atValue,
		};
	}

	const target = locatorToBlockTarget(locator);
	if (!target) {
		throw new EntryError(
			"INVALID_LOCATOR",
			`${labelPrefix}comments add anchors to a paragraph locator only (pN, pN:S-E, pN:S-pM:E, or tN:rRcC:pK[:S-E])`,
			"A whole table, section, comment, image, or other entity is not a valid anchor.",
		);
	}
	return {
		spec: {
			kind: "single",
			blockId: target.blockId,
			...(target.span ? { span: target.span } : {}),
		},
		text,
		author,
		locatorString: atValue,
	};
}

function resolveAnchorEntry(
	document: Document,
	anchor: string,
	occurrence: number,
	occurrenceExplicit: boolean,
	text: string,
	author: string,
	labelPrefix: string,
	findView: FindView,
): ResolvedEntry {
	const result = findTextSpans(document.body, anchor, { view: findView });
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
		spec: {
			kind: "single",
			blockId: match.blockId,
			span: { start: match.start, end: match.end },
		},
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
