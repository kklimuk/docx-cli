import {
	type BlockReference,
	type Document,
	Edit,
	EditError,
	getPageContentWidthEmu,
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
import { normalizeTabAlign, type TabStop } from "@core/blocks";
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
  (no --at)    PAGE SETUP for the WHOLE document: pass only --margins/
               --orientation/--size with no locator and EVERY section gets it.
               This is the "set it once for the whole document" path — each
               section has its own geometry, so on a multi-section doc this sets
               them all (incl. the trailing governing section that a per-section
               sweep tends to miss). Columns/type still need --at (which section?).
  --at sN      Edit an EXISTING section break sN in place — its columns, type,
               and/or page geometry. Use this to target ONE section's geometry;
               omit --at (above) to set the whole document. \`docx read FILE\`
               prints the section's id on the \`<!-- docx:page sN … -->\` note when
               geometry deviates from default.
  --at pN-pM   Wrap the paragraph range pN…pM in its own N-column section
               (inserts the bounding continuous breaks). Also accepts a single
               paragraph (--at pN). THIS is how you put text in columns — name the
               range. (Column wrap only; set page geometry on the resulting sN.)

Column / break options:
  --columns N       Number of columns (>= 1; use 1 to collapse back to single column)
  --type T          Section type: continuous (default), nextPage, evenPage,
                    oddPage, nextColumn. With a range it's the wrapping break's
                    type; with sN it overrides the section's type.

Page-setup options (no --at = whole document; --at sN = that one section):
  --orientation O   portrait | landscape (landscape swaps the page dimensions)
  --size SIZE       letter | legal | tabloid | a4 | a3 | a5, or WxH inches
                    (e.g. 8.5x11in, 8.5x14). A W>H size implies landscape.
  --margins M       One inch value (uniform, e.g. 1 or 1in) OR four comma-separated
                    values in CSS order top,right,bottom,left (e.g. 1,1,1,1.5).
                    Negative values are allowed (content into the margin).

  Changing margins/size auto-fixes tab columns: a template's right-edge LEFT tab
  (résumé dates/locations) calibrated to the old margins would overflow and WRAP
  at the new width, so page setup converts each to a RIGHT tab flush at the new
  margin (the \`--tabs right\` cure, applied for you). The ack reports how many were
  realigned. (Doc-wide, or a single-section doc; per-section edits on a
  multi-section doc don't reflow — \`read\`'s docx:layout hint flags those.)

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
  docx sections doc.docx --margins 0.5                  # 0.5in margins, WHOLE document (every section)
  docx sections doc.docx --orientation landscape        # whole-doc landscape (every section)
  docx sections doc.docx --margins 0.75,1,0.75,1        # top,right,bottom,left, whole document
  docx sections doc.docx --at p4-p9 --columns 2
  docx sections doc.docx --at p4-p9 --columns 3 --type continuous
  docx sections doc.docx --at s2 --columns 1            # collapse section s2 to one column
  docx sections doc.docx --at s2 --margins 0.5          # just section s2's margins
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

	const flags = await parseSectionFlags(parsed.values);
	if (typeof flags === "number") return flags;

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

	// No --at: page geometry is a DOCUMENT property, so `sections --margins/
	// --orientation/--size` (no locator) applies it to EVERY section. Each
	// `<w:sectPr>` carries its own pgMar/pgSz, so "uniform margins" on a
	// multi-section doc means setting them all — and the trailing governing
	// section is the easy one to miss (the resume s3 trap). Columns/type still
	// need a target (which section?), so reject those without --at.
	if (!at) {
		if (
			hasPageGeometry(flags) &&
			flags.columns === undefined &&
			flags.sectionType === undefined
		) {
			return editAllSections(document, opts);
		}
		return fail(
			"USAGE",
			"Missing --at LOCATOR (a section sN, or a range pN-pM to wrap in columns). To set PAGE GEOMETRY for the whole document, omit --at and pass only --margins/--orientation/--size.",
			HELP,
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
	// resolveTracked folds in the doc toggle — passing the raw `--track` boolean
	// (false) would short-circuit Edit.section's `track ?? toggle` and skip
	// tracking even when the document toggle is on.
	const track = resolveTracked(document, opts.trackFlag);
	try {
		new Edit(document).section(blockRef, opts.flags, {
			authorFlag: opts.authorFlag,
			track,
		});
	} catch (error) {
		if (error instanceof EditError) return fail(error.code, error.message);
		throw error;
	}
	// A single-section doc's only sectPr governs the whole body, so a geometry
	// edit here is effectively doc-wide — re-align fragile margin tabs too (the
	// editAllSections path does this for multi-section docs). For a targeted
	// section in a MULTI-section doc we don't reflow (we'd need per-section
	// paragraph governance); `read`'s docx:layout hint still surfaces it there.
	let realignedTabs = 0;
	if (hasPageGeometry(opts.flags) && countSections(document) === 1) {
		realignedTabs = reflowMarginTabs(document, {
			track,
			authorFlag: opts.authorFlag,
		});
	}
	return commit(document, opts, { at, mode: "section", realignedTabs });
}

/** Number of `<w:sectPr>` sections in the document (inline + trailing). */
function countSections(document: Document): number {
	let count = 0;
	for (const id of document.body.blockReferences.keys()) {
		if (/^s\d+$/.test(id)) count++;
	}
	return count;
}

/** Doc-wide page setup: apply the page geometry (margins/orientation/size) to
 * EVERY section in the document. Each `<w:sectPr>` has its own geometry, so this
 * is the "set it once for the whole document" affordance — it sets all of them
 * (including the trailing governing section a per-section sweep tends to miss).
 * Reuses the per-section `Edit.section` path, so under tracking each section
 * records its own `<w:sectPrChange>`. */
async function editAllSections(
	document: Document,
	opts: Options,
): Promise<number> {
	const sectionIds = [...document.body.blockReferences.keys()]
		.filter((id) => /^s\d+$/.test(id))
		.sort((left, right) => Number(left.slice(1)) - Number(right.slice(1)));
	if (sectionIds.length === 0) {
		return fail("BLOCK_NOT_FOUND", "No sections found to set page geometry on");
	}
	const track = resolveTracked(document, opts.trackFlag);
	try {
		const edit = new Edit(document);
		for (const id of sectionIds) {
			// Each section's <w:sectPr> mutates in place (no structural change), so
			// the other held refs stay valid across the loop.
			const blockRef = document.body.blockReferences.get(id);
			if (!blockRef) continue;
			edit.section(blockRef, opts.flags, {
				authorFlag: opts.authorFlag,
				track,
			});
		}
	} catch (error) {
		if (error instanceof EditError) return fail(error.code, error.message);
		throw error;
	}
	const realignedTabs = reflowMarginTabs(document, {
		track,
		authorFlag: opts.authorFlag,
	});
	return commit(document, opts, {
		at: `${sectionIds.length} section${sectionIds.length === 1 ? "" : "s"}`,
		mode: "section",
		realignedTabs,
	});
}

/** After a page-geometry change, re-align right-edge tab columns to the NEW text
 * margin — so we FIX the wrap instead of just warning about it. A template's
 * right-edge LEFT tab (the résumé date/location columns) is calibrated to the
 * ORIGINAL margins; widening the text area leaves it short of the new margin so
 * long values overflow and WRAP in render (the `docx:layout` hazard `read`
 * flags). Weak agents dismiss that hint as "informational" and ship the wrap, so
 * page setup applies the `--tabs right` cure itself: each fragile right-edge LEFT
 * tab becomes a RIGHT tab flush at the new margin. Returns the count realigned.
 *
 * Scope: body paragraphs only (cell tabs are table-relative, not page-margin),
 * non-list (a bullet's tab is the structural bullet-to-text gap), used (a real
 * tab run), right-edge (pos in the last ~30% of the text width), and not already
 * robust (no existing RIGHT tab). Keep this predicate in lockstep with
 * `layoutHazardNote` in `cli/read/markdown.ts` — flag and cure must agree. */
function reflowMarginTabs(
	document: Document,
	opts: { track: boolean; authorFlag?: string },
): number {
	const textWidthTwips = Math.round(getPageContentWidthEmu(document) / 635);
	const bodyChildren = document.body.findBodyChildren();
	const edit = new Edit(document);
	let realigned = 0;
	for (let index = 0; index < bodyChildren.length; index++) {
		const node = bodyChildren[index];
		if (!node || node.tag !== "w:p") continue;
		// A multi-column section's tab stops are column-relative, so the right-margin
		// cure doesn't apply (a right tab at the full text width lands outside the
		// narrower column) — `layoutHazardNote`'s `cols>1` branch deliberately
		// offers no cure, so skip those paragraphs here too.
		if (governingColumns(bodyChildren, index) > 1) continue;
		const tabs = reflowedTabStops(node, textWidthTwips);
		if (!tabs) continue;
		edit.paragraphProperties(
			{ node, parent: bodyChildren },
			{ tabs },
			{ authorFlag: opts.authorFlag, track: opts.track },
		);
		realigned++;
	}
	return realigned;
}

/** The paragraph's tab stops after curing fragile right-edge LEFT/center tabs,
 * or `null` when there's nothing to cure (so the caller skips it). PRESERVES every
 * non-fragile stop (a legit mid-line LEFT tab, a decimal/bar stop, …) and swaps
 * only the fragile right-edge one(s) for a SINGLE right tab flush at the new
 * margin — never the whole-paragraph `--tabs right` wipe (that dropped legit
 * stops). Mirrors `layoutHazardNote`'s single-column predicate in
 * `cli/read/markdown.ts` (right-edge, leftish, used, no existing right tab,
 * non-list); both classify a stop through `normalizeTabAlign`. */
function reflowedTabStops(
	node: XmlNode,
	textWidthTwips: number,
): TabStop[] | null {
	const pPr = node.findChild("w:pPr");
	if (!pPr) return null;
	if (pPr.findChild("w:numPr")) return null; // list — structural bullet tab
	const stops = (pPr.findChild("w:tabs")?.findChildren("w:tab") ?? [])
		.map((tab) => ({
			align: normalizeTabAlign(tab.getAttribute("w:val")),
			pos: Number(tab.getAttribute("w:pos") ?? "NaN"),
		}))
		.filter((stop) => Number.isFinite(stop.pos));
	if (stops.length === 0) return null;
	// An existing RIGHT tab already right-aligns robustly — leave the paragraph.
	if (stops.some((stop) => stop.align === "right")) return null;
	const isFragile = (stop: TabStop): boolean =>
		(stop.align === "left" || stop.align === "center") &&
		stop.pos > textWidthTwips * 0.7;
	if (!stops.some(isFragile)) return null;
	// Only when the tab is actually USED — a `<w:tab/>` CHARACTER in a run (the
	// thing that wraps), not just a stop definition. Keeps flag and cure aligned.
	if (!paragraphUsesTabCharacter(node)) return null;
	// Keep every non-fragile stop; replace the fragile right-edge one(s) with ONE
	// right tab at the margin. Sorted ascending (Word expects ordered stops).
	const merged: TabStop[] = [
		...stops.filter((stop) => !isFragile(stop)),
		{ align: "right", pos: textWidthTwips },
	];
	merged.sort((left, right) => left.pos - right.pos);
	return merged;
}

/** True if a run in the paragraph contains a `<w:tab/>` CHARACTER. Searches every
 * non-`<w:pPr>` subtree — inside `<w:pPr>` a `<w:tab>` is a STOP definition, not
 * a character, so it's skipped. */
function paragraphUsesTabCharacter(node: XmlNode): boolean {
	for (const child of node.children) {
		if (child.tag === "w:pPr") continue;
		if (containsTabCharacter(child)) return true;
	}
	return false;
}

function containsTabCharacter(node: XmlNode): boolean {
	if (node.tag === "w:tab") return true;
	return node.children.some(containsTabCharacter);
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
	meta: {
		at: string;
		mode: "section" | "range";
		dryRunOnly?: boolean;
		realignedTabs?: number;
	},
): Promise<number> {
	// `columnCount` rides along for --verbose/JSON consumers when columns were set
	// (a page-only edit has none). `locator` (not the count) is the salient field
	// the ack summarizer reports, so a page edit prints "sections s0", not a count.
	const columnCount =
		opts.flags.columns !== undefined ? { columnCount: opts.flags.columns } : {};
	// Surface the auto tab re-alignment so the agent KNOWS the wrap was cured (and
	// doesn't go re-fix it) — both in the JSON ack and the one-line confirmation.
	const realigned = meta.realignedTabs ?? 0;
	const realignedField = realigned > 0 ? { realignedTabs: realigned } : {};
	if (opts.dryRun || meta.dryRunOnly) {
		// Dry-run previews always print (no `ok` field), even without --verbose.
		await respond({
			operation: "sections",
			dryRun: true,
			path: opts.filePath,
			locator: meta.at,
			...columnCount,
			...realignedField,
			mode: meta.mode,
			...(opts.outputPath ? { output: opts.outputPath } : {}),
		});
		return EXIT.OK;
	}
	await document.save(opts.outputPath);
	const destination = opts.outputPath ?? opts.filePath;
	const realignedNote =
		realigned > 0
			? `re-aligned ${realigned} right-edge tab column${realigned === 1 ? "" : "s"} to the new margin (would have wrapped in render). `
			: "";
	await respondAck(
		{
			ok: true,
			operation: "sections",
			path: destination,
			locator: meta.at,
			...columnCount,
			...realignedField,
			mode: meta.mode,
		},
		realignedNote + renderVerifyHint(destination),
	);
	return EXIT.OK;
}
