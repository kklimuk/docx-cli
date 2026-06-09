import {
	type BlockReference,
	type Document,
	Edit,
	EditError,
	Insert,
	InsertError,
	isSectionType,
	type Locator,
	LocatorParseError,
	parseLocator,
	type SectionType,
	TrackChanges,
} from "@core";
import type { XmlNode } from "@core/parser";
import {
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
	resolveBlockRangeOrFail,
	resolveTracked,
	respond,
	respondAck,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx columns — lay text out in multiple columns

Usage:
  docx columns FILE --at LOCATOR --count N [options]

The intuitive verb for column layout, so you don't have to hand-build section
breaks. Two addressing modes:

  --at pN-pM   Wrap the paragraph range pN…pM in its own N-column section
               (inserts the bounding continuous section breaks for you). Also
               accepts a single paragraph (--at pN).
  --at sN      Set the column count on an EXISTING section break sN (the section
               whose content ENDS at sN). Equivalent to \`edit --at sN --columns N\`.

Options:
  --at LOCATOR      Paragraph range (pN-pM), single paragraph (pN), or section (sN)
  --count N         Number of columns (>= 1; use 1 to collapse back to single column)
  --type T          Section type for the wrapping break: continuous (default),
                    nextPage, evenPage, oddPage, nextColumn. Only meaningful with
                    a pN-pM/pN range; with sN it overrides the section's type.
  --author NAME     Author for tracked changes (default: $DOCX_AUTHOR)
  --track           Record as a tracked change even when the doc toggle is off
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -v, --verbose     Print the full success ack JSON
  -h, --help        Show this help

Agent tip: VERIFY LAYOUT VISUALLY. \`docx read\` shows the section as a
\`<!-- docx:section … cols="N" -->\` annotation but NOT how the columns actually
flow on the page (balance, overflow). After setting columns, render and look:
  docx render FILE --out pages/
Adjust the range and re-render until the columns read the way you intended.

Output:
  Silent on success (exit 0); --verbose prints {ok:true, …}. Positional ids shift
  after a range wrap (two section breaks are inserted), so re-read before further
  edits. Errors print {code, error, hint?} + nonzero exit.

Examples:
  docx columns doc.docx --at p4-p9 --count 2
  docx columns doc.docx --at p4-p9 --count 3 --type continuous
  docx columns doc.docx --at s2 --count 1        # collapse section s2 to one column
`;

const OPTION_SPEC = {
	at: { type: "string" },
	count: { type: "string" },
	type: { type: "string" },
	author: { type: "string" },
	track: { type: "boolean" },
	...SAVE_FLAGS,
} as const;

type Options = {
	filePath: string;
	count: number;
	sectionType?: SectionType;
	authorFlag?: string;
	trackFlag: boolean;
	outputPath?: string;
	dryRun: boolean;
};

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(args, OPTION_SPEC, HELP);
	if (typeof parsed === "number") return parsed;
	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}
	setVerboseAck(Boolean(parsed.values.verbose));

	const filePath = parsed.positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE argument", HELP);

	const at = parsed.values.at as string | undefined;
	if (!at) {
		return fail(
			"USAGE",
			"Missing --at LOCATOR (a range pN-pM or a section sN)",
			HELP,
		);
	}

	const count = parseCount(parsed.values.count as string | undefined);
	if (typeof count === "number" && Number.isNaN(count)) {
		return fail("USAGE", `--count must be a positive integer`, HELP);
	}
	if (count === undefined) return fail("USAGE", "Missing --count N", HELP);

	const typeRaw = parsed.values.type as string | undefined;
	if (typeRaw !== undefined && !isSectionType(typeRaw)) {
		return fail(
			"USAGE",
			`Invalid --type: ${typeRaw}`,
			"Valid values: continuous, nextPage, evenPage, oddPage, nextColumn",
		);
	}

	let locator: Locator;
	try {
		locator = parseLocator(at);
	} catch (error) {
		if (error instanceof LocatorParseError) {
			return fail("INVALID_LOCATOR", error.message, HELP);
		}
		throw error;
	}

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	const opts: Options = {
		filePath,
		count,
		sectionType: typeRaw,
		authorFlag: parsed.values.author as string | undefined,
		trackFlag: Boolean(parsed.values.track),
		outputPath: parsed.values.output as string | undefined,
		dryRun: Boolean(parsed.values["dry-run"]),
	};

	// pN-pM range, or a single pN paragraph, gets wrapped in a fresh section.
	// sN edits an existing section's column count in place.
	if (locator.kind === "blockRange") {
		const range = await resolveBlockRangeOrFail(document, at);
		if (typeof range === "number") return range;
		return wrapRange(
			document,
			range.parent,
			range.startIndex,
			range.endIndex,
			at,
			opts,
		);
	}

	const blockRef = await resolveBlockOrFail(document, at);
	if (typeof blockRef === "number") return blockRef;

	if (blockRef.node.tag === "w:sectPr") {
		return editSection(document, blockRef, at, opts);
	}
	if (blockRef.node.tag === "w:p") {
		const index = blockRef.parent.indexOf(blockRef.node);
		return wrapRange(document, blockRef.parent, index, index, at, opts);
	}
	return fail(
		"INVALID_LOCATOR",
		`columns needs a section (sN) or a paragraph range (pN-pM); ${at} is neither`,
		HELP,
	);
}

/** Set the column count on an existing section break — the thin-wrapper path,
 * identical to `edit --at sN --columns N`. */
async function editSection(
	document: Document,
	blockRef: BlockReference,
	at: string,
	opts: Options,
): Promise<number> {
	try {
		new Edit(document).section(
			blockRef,
			{ columns: opts.count, sectionType: opts.sectionType },
			{ authorFlag: opts.authorFlag, track: opts.trackFlag },
		);
	} catch (error) {
		if (error instanceof EditError) return fail(error.code, error.message);
		throw error;
	}
	return commit(document, opts, { at, mode: "section" });
}

/** Wrap [startIndex..endIndex] in its own N-column continuous section by
 * inserting the two bounding section breaks. The break AFTER the range carries
 * the requested column count; the break BEFORE the range preserves the
 * surrounding section's columns so the pre-range content is untouched. The
 * before-break is skipped when the range starts at the document's first block
 * (the new section then naturally begins at the top). */
async function wrapRange(
	document: Document,
	parent: XmlNode[],
	startIndex: number,
	endIndex: number,
	at: string,
	opts: Options,
): Promise<number> {
	// A `<w:sectPr>` may only live in `<w:body>` or a body-level paragraph's
	// `<w:pPr>` (ECMA-376) — never inside `<w:tc>`. Wrapping a cell paragraph
	// would write invalid OOXML ("unreadable content" in Word) AND be invisible
	// on read-back (the reader doesn't enumerate in-cell sectPr), silently
	// breaking the write-read loop. Reject anything not in the body.
	if (parent !== document.body.findBodyChildren()) {
		return fail(
			"USAGE",
			`columns can only wrap body-level paragraphs; ${at} is inside a table cell`,
			"Section breaks carry column layout and cannot live in a table cell.",
		);
	}

	for (let index = startIndex; index <= endIndex; index++) {
		if (containsSectionBreak(parent[index])) {
			return fail(
				"USAGE",
				`The range ${at} already contains a section break`,
				"Target the existing section directly with `--at sN`, or pick a range without one.",
			);
		}
	}

	const governing = governingColumns(parent, endIndex + 1);
	const track = resolveTracked(document, opts.trackFlag);
	const insert = new Insert(document);
	// Shared allocator so the two tracked inserts don't mint duplicate revision
	// ids (neither block is spliced when the other is built, so per-call
	// allocators would both see the same tree max).
	const allocator = track
		? new TrackChanges(document).createAllocator()
		: undefined;

	const endNode = parent[endIndex];
	if (!endNode) {
		return fail(
			"BLOCK_NOT_FOUND",
			"Range end is stale (parent does not contain it)",
		);
	}
	const startNode = startIndex > 0 ? parent[startIndex] : undefined;
	if (startIndex > 0 && !startNode) {
		return fail(
			"BLOCK_NOT_FOUND",
			"Range start is stale (parent does not contain it)",
		);
	}

	let afterBlocks: XmlNode[];
	let beforeBlocks: XmlNode[] = [];
	try {
		afterBlocks = await insert.paragraph(
			{ node: endNode, parent },
			{
				kind: "section",
				columns: opts.count,
				sectionType: opts.sectionType ?? "continuous",
			},
			{},
			{ placement: "after", authorFlag: opts.authorFlag, track, allocator },
		);
		if (startNode) {
			beforeBlocks = await insert.paragraph(
				{ node: startNode, parent },
				{
					kind: "section",
					...(governing > 1 ? { columns: governing } : {}),
					sectionType: "continuous",
				},
				{},
				{ placement: "before", authorFlag: opts.authorFlag, track, allocator },
			);
		}
	} catch (error) {
		if (error instanceof InsertError)
			return fail(error.code, error.message, error.hint);
		throw error;
	}

	if (opts.dryRun) {
		return commit(document, opts, { at, mode: "range", dryRunOnly: true });
	}

	// Splice the after-break first (higher index) so the before-break's index
	// stays valid.
	parent.splice(endIndex + 1, 0, ...afterBlocks);
	parent.splice(startIndex, 0, ...beforeBlocks);
	return commit(document, opts, { at, mode: "range" });
}

/** True if `node` is a section break — the trailing `<w:sectPr>` or a paragraph
 * carrying an inline `<w:pPr><w:sectPr>`. */
function containsSectionBreak(node: XmlNode | undefined): boolean {
	if (!node) return false;
	if (node.tag === "w:sectPr") return true;
	if (node.tag !== "w:p") return false;
	return Boolean(node.findChild("w:pPr")?.findChild("w:sectPr"));
}

/** The column count of the section governing the block at/after `fromIndex` —
 * the first section break encountered scanning forward (an inline sectPr, or the
 * trailing one). Defaults to 1 (single column) when none declares `<w:cols>`. */
function governingColumns(parent: XmlNode[], fromIndex: number): number {
	for (let index = fromIndex; index < parent.length; index++) {
		const node = parent[index];
		if (!node) continue;
		if (node.tag === "w:sectPr") return columnsOf(node);
		if (node.tag === "w:p") {
			const inline = node.findChild("w:pPr")?.findChild("w:sectPr");
			if (inline) return columnsOf(inline);
		}
	}
	return 1;
}

function columnsOf(sectPr: XmlNode): number {
	const raw = sectPr.findChild("w:cols")?.getAttribute("w:num");
	const parsed = raw ? Number.parseInt(raw, 10) : 1;
	return Number.isFinite(parsed) && parsed > 0 ? parsed : 1;
}

/** Parse `--count`. Returns undefined when absent, NaN-number on invalid, or
 * the positive integer. */
function parseCount(raw: string | undefined): number | undefined {
	if (raw === undefined) return undefined;
	// Require pure digits — bare `Number.parseInt` would silently TRUNCATE
	// "2.5" → 2 and "1e2" → 1, both of which then pass the `>= 1` check and
	// write a wrong column count with no error.
	if (!/^\d+$/.test(raw.trim())) return Number.NaN;
	const parsed = Number.parseInt(raw, 10);
	if (!Number.isFinite(parsed) || parsed < 1) return Number.NaN;
	return parsed;
}

async function commit(
	document: Document,
	opts: Options,
	meta: { at: string; mode: "section" | "range"; dryRunOnly?: boolean },
): Promise<number> {
	if (opts.dryRun || meta.dryRunOnly) {
		// Dry-run previews always print (no `ok` field), even without --verbose.
		await respond({
			operation: "columns",
			dryRun: true,
			path: opts.filePath,
			at: meta.at,
			count: opts.count,
			mode: meta.mode,
			...(opts.outputPath ? { output: opts.outputPath } : {}),
		});
		return EXIT.OK;
	}
	await document.save(opts.outputPath);
	await respondAck({
		ok: true,
		operation: "columns",
		path: opts.outputPath ?? opts.filePath,
		at: meta.at,
		count: opts.count,
		mode: meta.mode,
	});
	return EXIT.OK;
}
