import {
	type BlockReference,
	type Document,
	Edit,
	EditError,
	Insert,
	InsertError,
	inheritPageGeometry,
	type Locator,
	LocatorParseError,
	type PageGeometry,
	parseLocator,
	type SectionType,
	TrackChanges,
} from "@core";
import type { XmlNode } from "@core/parser";
import { parseSectionFlags } from "../parse-helpers";
import {
	EXIT,
	fail,
	openOrFail,
	renderVerifyHint,
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

/** The parsed section flags (columns/type + page geometry) plus a derived flag
 *  for "any page geometry was requested" — page geometry is only valid on an
 *  existing section (sN), never on a range/paragraph column wrap. */
type SectionFlags = {
	columns?: number;
	sectionType?: SectionType;
} & PageGeometry;

function hasPageGeometry(flags: SectionFlags): boolean {
	return (
		flags.pageSize !== undefined ||
		flags.orientation !== undefined ||
		flags.margins !== undefined
	);
}

const HELP = `docx sections — multi-column layout, section breaks & page setup

Usage:
  docx sections FILE --at LOCATOR (--columns N | page setup) [options]

The verb for everything that lives in a \`<w:sectPr>\` — multi-column flow,
section breaks, AND page geometry (margins / orientation / size) — so you don't
have to hand-build them with the right OOXML semantics. This is the ONLY way to
add/change column layout; \`insert\` no longer takes --section (a raw section
break formats the content ABOVE it, which is the classic off-by-one).

Addressing:
  --at sN      Edit an EXISTING section break sN in place — its columns, type,
               and/or page geometry. PAGE SETUP applies to the whole document via
               its trailing section break (the LAST sN — s0 in a single-section
               doc; \`docx read FILE\` prints the section's id on the
               \`<!-- docx:page sN … -->\` note when geometry deviates from default).
  --at pN-pM   Wrap the paragraph range pN…pM in its own N-column section
               (inserts the bounding continuous breaks). Also accepts a single
               paragraph (--at pN). THIS is how you put text in columns — name the
               range. (Column wrap only; set page geometry on the resulting sN.)

Column / break options:
  --columns N       Number of columns (>= 1; use 1 to collapse back to single column)
  --type T          Section type: continuous (default), nextPage, evenPage,
                    oddPage, nextColumn. With a range it's the wrapping break's
                    type; with sN it overrides the section's type.

Page-setup options (with --at sN — set the document/section page geometry in place):
  --orientation O   portrait | landscape (landscape swaps the page dimensions)
  --size SIZE       letter | legal | tabloid | a4 | a3 | a5, or WxH inches
                    (e.g. 8.5x11in, 8.5x14). A W>H size implies landscape.
  --margins M       One inch value (uniform, e.g. 1 or 1in) OR four comma-separated
                    values in CSS order top,right,bottom,left (e.g. 1,1,1,1.5).
                    Negative values are allowed (content into the margin).

General:
  --at LOCATOR      Section (sN), paragraph range (pN-pM), or single paragraph (pN)
  --author NAME     Author for tracked changes (default: $DOCX_AUTHOR)
  --track           Record as a tracked change even when the doc toggle is off
                    (page/column/type edits record as ONE <w:sectPrChange>)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -v, --verbose     Print the full success ack JSON
  -h, --help        Show this help

Agent tip: VERIFY LAYOUT VISUALLY. \`docx read\` shows columns as a
\`<!-- docx:section … cols="N" -->\` annotation and page geometry as a leading
\`<!-- docx:page sN orientation=… margins=… -->\` note, but NOT how content flows
on the page. After a layout change, render and look:
  docx render FILE --out pages/

Output:
  Prints a one-line confirmation on success (exit 0); --verbose prints {ok:true, …}. Positional ids shift
  after a range wrap (two section breaks are inserted), so re-read before further
  edits. Errors print {code, error, hint?} + nonzero exit.

Examples:
  docx sections doc.docx --at p4-p9 --columns 2
  docx sections doc.docx --at p4-p9 --columns 3 --type continuous
  docx sections doc.docx --at s2 --columns 1            # collapse section s2 to one column
  docx sections doc.docx --at s0 --orientation landscape # whole-doc landscape (trailing sN)
  docx sections doc.docx --at s0 --margins 1in --size letter
  docx sections doc.docx --at s0 --margins 0.75,1,0.75,1 # top,right,bottom,left
`;

const OPTION_SPEC = {
	at: { type: "string" },
	columns: { type: "string" },
	type: { type: "string" },
	orientation: { type: "string" },
	size: { type: "string" },
	margins: { type: "string" },
	author: { type: "string" },
	track: { type: "boolean" },
	...SAVE_FLAGS,
} as const;

type Options = {
	filePath: string;
	flags: SectionFlags;
	authorFlag?: string;
	trackFlag: boolean;
	outputPath?: string;
	dryRun: boolean;
};

// Wrapping a range/paragraph in a fresh section is a COLUMN operation — page
// geometry has no meaning there (it applies to an existing sectPr), so reject it
// with a pointer to the right move.
const WRAP_REJECTS_GEOMETRY =
	"--orientation/--size/--margins set page geometry on an EXISTING section (--at sN); they don't apply to a column wrap. Wrap the range first, then `docx sections --at sN --orientation/--size/--margins …` (the trailing sN is the whole document).";

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
			"Missing --at LOCATOR (a section sN, or a range pN-pM to wrap in columns)",
			HELP,
		);
	}

	const flags = await parseSectionFlags(parsed.values);
	if (typeof flags === "number") return flags;

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
		flags,
		authorFlag: parsed.values.author as string | undefined,
		trackFlag: Boolean(parsed.values.track),
		outputPath: parsed.values.output as string | undefined,
		dryRun: Boolean(parsed.values["dry-run"]),
	};

	// pN-pM range, or a single pN paragraph, gets wrapped in a fresh column
	// section (requires --columns; page geometry not allowed). sN edits an
	// existing section in place — columns, type, and/or page geometry.
	if (locator.kind === "blockRange") {
		if (hasPageGeometry(flags))
			return fail("USAGE", WRAP_REJECTS_GEOMETRY, HELP);
		if (flags.columns === undefined) {
			return fail(
				"USAGE",
				"Wrapping a range in a section needs --columns N",
				HELP,
			);
		}
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
		if (
			flags.columns === undefined &&
			flags.sectionType === undefined &&
			!hasPageGeometry(flags)
		) {
			return fail(
				"USAGE",
				"Section edit needs --columns, --type, --orientation, --size, or --margins",
				HELP,
			);
		}
		return editSection(document, blockRef, at, opts);
	}
	if (blockRef.node.tag === "w:p") {
		if (hasPageGeometry(flags))
			return fail("USAGE", WRAP_REJECTS_GEOMETRY, HELP);
		if (flags.columns === undefined) {
			return fail(
				"USAGE",
				"Wrapping a paragraph in a section needs --columns N",
				HELP,
			);
		}
		const index = blockRef.parent.indexOf(blockRef.node);
		return wrapRange(document, blockRef.parent, index, index, at, opts);
	}
	return fail(
		"INVALID_LOCATOR",
		`--at needs a section (sN) or a paragraph range (pN-pM); ${at} is neither`,
		HELP,
	);
}

/** Edit an existing section break in place — columns, type, and/or page geometry
 * (margins/orientation/size). The thin-wrapper path, identical to
 * `edit --at sN …` for columns/type; page geometry rides the same Edit.section. */
async function editSection(
	document: Document,
	blockRef: BlockReference,
	at: string,
	opts: Options,
): Promise<number> {
	try {
		new Edit(document).section(blockRef, opts.flags, {
			authorFlag: opts.authorFlag,
			// resolveTracked folds in the doc toggle — passing the raw `--track`
			// boolean (false) would short-circuit Edit.section's `track ?? toggle`
			// and skip tracking even when the document toggle is on.
			track: resolveTracked(document, opts.trackFlag),
		});
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
	// The section currently governing the wrapped range — its page geometry
	// (size/orientation/margins) must flow into the fresh sentinel sectPrs, or
	// the wrap silently reverts those new sections to portrait-Letter (the
	// "landscape vanishes after adding columns" footgun).
	const governingGeometry = governingSectPr(parent, endIndex + 1);
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
				columns: opts.flags.columns,
				sectionType: opts.flags.sectionType ?? "continuous",
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

	// Carry the governing section's page geometry onto every fresh sentinel sectPr
	// so the wrap preserves the document's size/orientation/margins (each section's
	// geometry is independent — without this the new sections revert to portrait).
	if (governingGeometry) {
		for (const block of [...afterBlocks, ...beforeBlocks]) {
			const sectPr = block.findChild("w:pPr")?.findChild("w:sectPr");
			if (sectPr) inheritPageGeometry(sectPr, governingGeometry);
		}
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

/** The `<w:sectPr>` governing the block at/after `fromIndex` — the first section
 * break scanning forward (an inline sectPr in a paragraph's pPr, or the trailing
 * body sectPr). Its page geometry is what a fresh column-wrap section inherits.
 * Returns undefined only if no sectPr follows (shouldn't happen — the trailing
 * one is mandatory). */
function governingSectPr(
	parent: XmlNode[],
	fromIndex: number,
): XmlNode | undefined {
	for (let index = fromIndex; index < parent.length; index++) {
		const node = parent[index];
		if (!node) continue;
		if (node.tag === "w:sectPr") return node;
		if (node.tag === "w:p") {
			const inline = node.findChild("w:pPr")?.findChild("w:sectPr");
			if (inline) return inline;
		}
	}
	return undefined;
}

async function commit(
	document: Document,
	opts: Options,
	meta: { at: string; mode: "section" | "range"; dryRunOnly?: boolean },
): Promise<number> {
	// `columnCount` rides along for --verbose/JSON consumers when columns were set
	// (a page-only edit has none). `locator` (not the count) is the salient field
	// the ack summarizer reports, so a page edit prints "sections s0", not a count.
	const columnCount =
		opts.flags.columns !== undefined ? { columnCount: opts.flags.columns } : {};
	if (opts.dryRun || meta.dryRunOnly) {
		// Dry-run previews always print (no `ok` field), even without --verbose.
		await respond({
			operation: "sections",
			dryRun: true,
			path: opts.filePath,
			locator: meta.at,
			...columnCount,
			mode: meta.mode,
			...(opts.outputPath ? { output: opts.outputPath } : {}),
		});
		return EXIT.OK;
	}
	await document.save(opts.outputPath);
	const destination = opts.outputPath ?? opts.filePath;
	await respondAck(
		{
			ok: true,
			operation: "sections",
			path: destination,
			locator: meta.at,
			...columnCount,
			mode: meta.mode,
		},
		renderVerifyHint(destination),
	);
	return EXIT.OK;
}
