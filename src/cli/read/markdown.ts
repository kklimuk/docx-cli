import {
	type Block,
	type Body,
	type Comment,
	type Footnote,
	flattenParagraphs,
	type ImageRun,
	type Locator,
	LocatorParseError,
	type Marginal,
	type Paragraph,
	parseLocator,
	type Run,
	type SectionBreak,
	type Table,
	type TableCell,
	type TableRow,
	type TextRun,
	type TrackedChange,
} from "@core";
import {
	codeBlockLanguageFromStyleId,
	isCodeBlockStyleId,
} from "@core/code-block";
import { extensionForImageMime } from "@core/image/formats";
import { inlineEscapeMask } from "@core/markdown";
import {
	emuToInches,
	formatNote,
	htmlAttr,
	type NotePair,
	twipsToInches,
	twipsToPoints,
} from "./annotations";

export type MarkdownView = "current" | "accepted" | "baseline";

export type MarkdownOptions = {
	from?: string;
	to?: string;
	view?: MarkdownView;
	showComments?: boolean;
	/** Document default run size in half-points (from styles `docDefaults`).
	 *  Used as the baseline size when no stronger per-run majority emerges, so a
	 *  doc that stamps the default `<w:sz>` on most runs still reads clean. */
	defaultSizeHalfPoints?: number;
	/** Document default run font (explicit `w:ascii` from styles `docDefaults`).
	 *  Surfaced in the `docx:base` note when it DEVIATES from the template default
	 *  (i.e. `set-default-font` ran), so the document font is observable on read. */
	defaultFont?: string;
	/** Whether the document's global track-changes toggle is on. Surfaced as a
	 *  head `<!-- docx:track-changes on -->` hint (deviation-only — off is the
	 *  default, so nothing is emitted then) so an agent sees the state on the
	 *  first read instead of spending a `track-changes list` call to learn it. */
	trackChangesOn?: boolean;
};

/** Document-wide run formatting so ubiquitous it reads as noise: the dominant
 *  font and size across all body + table runs. `read` emits it once as a
 *  `<!-- docx:base … -->` note and omits it from every matching run, so the body
 *  reads clean AND an agent can see the doc's baseline to match new content. It's
 *  a VISIBILITY hint (per "comments are never anything but hints") — the importer
 *  DROPS it, so a full `read → create` rebuild falls back to the template
 *  docDefaults for the dominant font/size (`read --ast` stays lossless;
 *  in-place `edit` preserves runs). A value only becomes a baseline when it
 *  covers a majority of the document's text — otherwise every run keeps its
 *  explicit formatting. */
export type RunFormatBaseline = { font?: string; sizeHalfPoints?: number };

/** The canonical create-template `docDefaults` (Calibri 11pt) — the universal
 *  default, noise on every doc, so suppressed from the `docx:base` note. A
 *  docDefaults value that DEVIATES from it is the meaningful signal that
 *  `set-default-font` ran, and IS surfaced. */
const TEMPLATE_DEFAULT_FONT = "Calibri";
const TEMPLATE_DEFAULT_SIZE_HALF_POINTS = 22;

/** Return `value` only when it deviates from the canonical template `def` (else
 *  undefined): the universal default is noise; a deviation is the signal. */
function deviation<T>(value: T | undefined, def: T): T | undefined {
	return value !== undefined && value !== def ? value : undefined;
}

export class MarkdownLocatorError extends Error {
	constructor(
		public input: string,
		message: string,
	) {
		super(message);
		this.name = "MarkdownLocatorError";
	}
}

type CommentIndex = {
	endingsByRun: Map<string, string[]>;
	spanText: Map<string, string>;
	orderedIds: string[];
};

type RenderContext = {
	options: MarkdownOptions;
	baseline: RunFormatBaseline;
	commentIndex: CommentIndex;
	referencedFootnoteIds: Set<string>;
	referencedEndnoteIds: Set<string>;
	referencedTrackedChanges: Map<string, TrackedChange>;
	/** Running count of ordered-list items seen, keyed by `${numId}:${level}`, so
	 *  each item renders its real ordinal (1. 2. 3.) instead of a uniform `1.` —
	 *  the raw markdown reads correctly, not as a wall of `1.` (a renderer would
	 *  auto-increment, but the raw text an agent/human reads would not). */
	orderedCounters: Map<string, number>;
	/** Usable page width in EMU (page width − L/R margins) from the document's
	 *  geometry, for flagging images wider than the text column (`overflow`). */
	contentWidthEmu: number;
	/** Per-block governing column count (the columns of the section break the
	 *  block falls under). Used to flag tab-aligned content living in a
	 *  multi-column section, where tab stops wrap mid-line in the narrow column —
	 *  a render-only break Markdown can't show (the résumé scenario's blocker). */
	governingColumns: Map<string, number>;
	/** Paragraph ids whose trailing tab-aligned content WILL wrap in render (a
	 *  right-edge LEFT tab). Collected by `layoutHazardNote` during the walk so
	 *  `renderMarkdown` can emit one consolidated summary + single-call cure. */
	wrappingTabLines: string[];
	/** Per-section `docx:header`/`docx:footer` hint lines, keyed by section id, for
	 *  marginals whose text DIFFERS by section. Co-located with each section break
	 *  in the body (the trailing section's land at the end) so a multi-section doc's
	 *  header/footer hints sit next to the section they govern, not bunched at the
	 *  top. UNIFORM marginals (same across every section) stay in the head block. */
	marginalsBySection: Map<string, string[]>;
};

export function renderMarkdown(
	doc: Body,
	options: MarkdownOptions = {},
): string {
	const blocks = sliceBlocks(doc.blocks, options.from, options.to);
	const dominant = detectFormatBaseline(blocks);
	// Suppression falls back to the document default size, so a doc that stamps the
	// default `<w:sz>` on most runs still reads clean.
	const baseline: RunFormatBaseline = {
		font: dominant.font,
		sizeHalfPoints: dominant.sizeHalfPoints ?? options.defaultSizeHalfPoints,
	};
	// The NOTE declares the document's baseline font/size: the per-run majority,
	// else the document DEFAULT (docDefaults) when it DEVIATES from the canonical
	// template (Calibri 11pt) — i.e. someone ran `set-default-font`. The universal
	// template default is suppressed as noise; a deviation is the meaningful signal
	// that makes the document font/size observable on read. Safe to declare: the
	// note is a hint the importer drops, so it can't drift a `read → create` rebuild.
	const noteBaseline: RunFormatBaseline = {
		font:
			dominant.font ?? deviation(options.defaultFont, TEMPLATE_DEFAULT_FONT),
		sizeHalfPoints:
			dominant.sizeHalfPoints ??
			deviation(
				options.defaultSizeHalfPoints,
				TEMPLATE_DEFAULT_SIZE_HALF_POINTS,
			),
	};
	const commentIndex = options.showComments
		? buildCommentIndex(blocks, options)
		: emptyCommentIndex();
	// Header/footer hints split two ways: UNIFORM marginals (identical across
	// every section — document chrome) ride the head block; PER-SECTION marginals
	// (the header/footer differs by section) co-locate with their section break in
	// the body (the trailing section's land at the end), so the agent sees each
	// next to the section it governs instead of bunched at the top.
	const sectionCount = blocks.filter(
		(block) => block.type === "sectionBreak",
	).length;
	const marginal = formatMarginalNotes(doc.headers, doc.footers, sectionCount);
	const ctx: RenderContext = {
		options,
		baseline,
		commentIndex,
		referencedFootnoteIds: new Set(),
		referencedEndnoteIds: new Set(),
		referencedTrackedChanges: new Map(),
		orderedCounters: new Map(),
		contentWidthEmu: contentWidthEmu(documentGeometry(blocks)?.geometry),
		governingColumns: computeGoverningColumns(blocks),
		wrappingTabLines: [],
		marginalsBySection: marginal.bySection,
	};

	// Sections render at their START (where their content begins), not at the
	// sectPr's physical end, so each section's `docx:section` + header/footer
	// annotation reads right before the content it governs (`applies-to="… (below)"`).
	// This is a read-view choice — the AST and `sN` addressing are unchanged.
	const sectionsByStart = computeSectionStarts(blocks);
	const parts: string[] = [];
	let cursor = 0;
	while (cursor < blocks.length) {
		// Emit the annotation of every section whose content STARTS here, before the
		// section's first block.
		for (const entry of sectionsByStart.get(cursor) ?? []) {
			const startNote = renderSectionStart(entry, blocks, ctx);
			if (startNote) parts.push(startNote);
		}
		const block = blocks[cursor];
		if (!block) {
			cursor++;
			continue;
		}
		// Section breaks no longer render in place — their annotation was emitted at
		// the section's start (above).
		if (block.type === "sectionBreak") {
			cursor++;
			continue;
		}
		// Collapse a run of CodeBlock paragraphs into one fenced GFM block.
		// Walk forward as long as adjacent blocks are CodeBlock paragraphs;
		// emit one ```...``` for the group. Locator comments go on their own
		// lines bracketing the fence (not on the fence lines, which would corrupt
		// the language id and break fence-closing), so the block reads cleanly.
		if (isCodeBlockParagraph(block)) {
			let lookahead = cursor + 1;
			while (lookahead < blocks.length) {
				const next = blocks[lookahead];
				if (!next || !isCodeBlockParagraph(next)) break;
				lookahead++;
			}
			const group = blocks.slice(cursor, lookahead) as Paragraph[];
			parts.push(renderCodeBlockGroup(group, ctx));
			cursor = lookahead;
			continue;
		}
		const rendered = renderBlock(block, ctx);
		if (rendered !== null) parts.push(rendered);
		cursor++;
	}
	const definitions: string[] = [];
	if (options.showComments) {
		const commentFootnotes = renderCommentFootnotes(commentIndex, doc.comments);
		if (commentFootnotes.length > 0) definitions.push(commentFootnotes);
	}
	const footnoteDefs = renderNoteDefinitions(
		doc.footnotes,
		ctx.referencedFootnoteIds,
	);
	if (footnoteDefs.length > 0) definitions.push(footnoteDefs);
	const endnoteDefs = renderNoteDefinitions(
		doc.endnotes,
		ctx.referencedEndnoteIds,
	);
	if (endnoteDefs.length > 0) definitions.push(endnoteDefs);
	const trackedChangeDefs = renderTrackedChangeFootnotes(
		ctx.referencedTrackedChanges,
	);
	if (trackedChangeDefs.length > 0) definitions.push(trackedChangeDefs);
	if (definitions.length > 0) parts.push(definitions.join("\n"));
	if (parts.length === 0) return "";
	// The base note rides at the very top so the importer reads it before any
	// content — it declares the document's dominant font/size, which `read` then
	// omits from every matching run below. The page note follows it, declaring
	// document page geometry when it deviates from the canonical default (a
	// read-time hint the importer drops; geometry survives edits in place).
	const baseNote = formatBaseNote(noteBaseline);
	const pageNote = formatPageNote(documentGeometry(blocks));
	// Track-changes ON is a deviation from the default (off), so surface it; off
	// emits nothing. This is an own-`docx:` metadata hint, dropped on import like
	// the others — it just orients the agent so edits/`--track` decisions don't
	// need a separate `track-changes list`.
	const trackNote = options.trackChangesOn
		? formatNote("track-changes", [], ["on"])
		: "";
	// One consolidated, actionable summary for every line that WILL wrap in render —
	// the per-line hints alone got ignored by both Haiku and Sonnet across repeated
	// reads, so lead with a single command that cures them all.
	const wrapSummary = formatWrappingTabSummary(ctx.wrappingTabLines);
	// Only the UNIFORM header/footer hints (identical across every section) ride
	// the head block next to the page note — document chrome. Per-section ones
	// were already emitted next to their section in the body loop above. Content
	// lives in the comment attribute, so the importer's blanket `docx:` drop keeps
	// a read → create rebuild from duplicating it into the body.
	const headLines = [
		baseNote,
		pageNote,
		...marginal.uniform,
		trackNote,
		wrapSummary,
	].filter((line) => line.length > 0);
	const head = headLines.length > 0 ? `${headLines.join("\n")}\n\n` : "";
	return `${head}${parts.join("\n\n")}\n`;
}

/** The dominant font and size across all body + table runs, each reported only
 * when it covers a majority of the document's rendered text (weighted by
 * character count). A clear majority is what makes omitting it from every run
 * legible rather than lossy; below the threshold the document has no single
 * baseline and every run keeps its explicit formatting. */
function detectFormatBaseline(blocks: Block[]): RunFormatBaseline {
	const fontChars = new Map<string, number>();
	const sizeChars = new Map<number, number>();
	let total = 0;
	for (const paragraph of flattenParagraphs(blocks)) {
		for (const run of paragraph.runs) {
			if (run.type !== "text") continue;
			const length = run.text.length;
			if (length === 0) continue;
			total += length;
			if (run.font)
				fontChars.set(run.font, (fontChars.get(run.font) ?? 0) + length);
			if (run.sizeHalfPoints !== undefined) {
				sizeChars.set(
					run.sizeHalfPoints,
					(sizeChars.get(run.sizeHalfPoints) ?? 0) + length,
				);
			}
		}
	}
	if (total === 0) return {};
	return {
		font: majorityKey(fontChars, total),
		sizeHalfPoints: majorityKey(sizeChars, total),
	};
}

/** The map key whose accumulated weight exceeds half the total, or undefined
 * when no single key does. */
function majorityKey<K>(counts: Map<K, number>, total: number): K | undefined {
	for (const [key, weight] of counts) {
		if (weight * 2 > total) return key;
	}
	return undefined;
}

/** The `<!-- docx:base … -->` line, or "" when there's no baseline to declare —
 * `size` in points, `font` verbatim. A visibility hint the importer drops (not
 * parse-back). `formatNote` applies the shared `htmlAttr` escaping so a
 * pathological font value (a `"`, `>`, or `-->`) can't break the comment. */
function formatBaseNote(baseline: RunFormatBaseline): string {
	const pairs: NotePair[] = [];
	if (baseline.font) pairs.push(["font", baseline.font]);
	if (baseline.sizeHalfPoints !== undefined) {
		pairs.push(["size", `${baseline.sizeHalfPoints / 2}pt`]);
	}
	return pairs.length > 0 ? formatNote("base", pairs) : "";
}

function emptyCommentIndex(): CommentIndex {
	return {
		endingsByRun: new Map(),
		spanText: new Map(),
		orderedIds: [],
	};
}

function isRunVisible(run: TextRun, view: MarkdownView): boolean {
	const kind = run.trackedChange?.kind;
	if (!kind) return true;
	if (view === "accepted" && (kind === "del" || kind === "moveFrom"))
		return false;
	if (view === "baseline" && (kind === "ins" || kind === "moveTo"))
		return false;
	return true;
}

function buildCommentIndex(
	blocks: Block[],
	options: MarkdownOptions,
): CommentIndex {
	const view = options.view ?? "accepted";
	const lastSlot = new Map<string, string>();
	const spanText = new Map<string, string>();
	const orderedIds: string[] = [];

	for (const paragraph of flattenParagraphs(blocks)) {
		paragraph.runs.forEach((run, index) => {
			// Text runs are the primary carrier; equation runs also need to
			// pick up comments (audit-comment fallback for tracked equation
			// edits anchors comment ranges on the `<m:oMath>` itself).
			const comments = runComments(run);
			if (!comments) return;
			if (run.type === "text" && !isRunVisible(run, view)) return;
			const spanContribution =
				run.type === "text"
					? run.text
					: run.type === "equation"
						? run.text
						: "";
			for (const commentId of comments) {
				if (!spanText.has(commentId)) orderedIds.push(commentId);
				spanText.set(
					commentId,
					(spanText.get(commentId) ?? "") + spanContribution,
				);
				lastSlot.set(commentId, slotKey(paragraph.id, index));
			}
		});
	}

	const endingsByRun = new Map<string, string[]>();
	for (const commentId of orderedIds) {
		const slot = lastSlot.get(commentId);
		if (!slot) continue;
		const list = endingsByRun.get(slot) ?? [];
		list.push(commentId);
		endingsByRun.set(slot, list);
	}

	return { endingsByRun, spanText, orderedIds };
}

/** Comment IDs attached to a run, regardless of run type. Today text runs
 *  and equation runs are the only carriers — extend here if other run types
 *  start carrying comment anchors. */
function runComments(run: Run): string[] | undefined {
	if (run.type === "text") return run.comments;
	if (run.type === "equation") return run.comments;
	return undefined;
}

function slotKey(paragraphId: string, runIndex: number): string {
	return `${paragraphId}#${runIndex}`;
}

function renderBlock(block: Block, ctx: RenderContext): string | null {
	if (block.type === "paragraph") return renderParagraph(block, ctx);
	if (block.type === "table") return renderTable(block, ctx);
	// Section breaks render at their section's start (see `renderSectionStart`),
	// not here.
	return null;
}

type SectionStart = {
	block: SectionBreak;
	breakIndex: number;
};

/** Map each section's CONTENT-START index to the section(s) beginning there. A
 *  section spans from the block after the previous break (or 0) through its own
 *  break; it STARTS at that first block. The trailing mandatory sectPr is the
 *  last block. */
function computeSectionStarts(blocks: Block[]): Map<number, SectionStart[]> {
	const starts = new Map<number, SectionStart[]>();
	let sectionStart = 0;
	for (let index = 0; index < blocks.length; index++) {
		if (blocks[index]?.type !== "sectionBreak") continue;
		const entry: SectionStart = {
			block: blocks[index] as SectionBreak,
			breakIndex: index,
		};
		const group = starts.get(sectionStart) ?? [];
		group.push(entry);
		starts.set(sectionStart, group);
		sectionStart = index + 1;
	}
	return starts;
}

/** Render a section's annotation at its content START: the `docx:section` note
 *  (cols/type deviation-only, `applies-to="… (below)"`) plus its per-section
 *  header/footer hints. EVERY section emits its marker — the trailing mandatory
 *  sectPr included — so the read consistently shows where each section begins. */
function renderSectionStart(
	entry: SectionStart,
	blocks: Block[],
	ctx: RenderContext,
): string {
	return renderSectionBreak(
		entry.block,
		governedRange(blocks, entry.breakIndex),
		ctx.marginalsBySection.get(entry.block.id),
	);
}

/** Map each block id to the column count of the section that governs it (the
 * first section break at or after it — sections govern preceding content). Built
 * by scanning from the end: the running column count is the nearest section
 * break already passed. Default 1 (single column). */
function computeGoverningColumns(blocks: Block[]): Map<string, number> {
	const map = new Map<string, number>();
	let cols = 1;
	for (let index = blocks.length - 1; index >= 0; index--) {
		const block = blocks[index];
		if (!block) continue;
		if (block.type === "sectionBreak") {
			cols = block.columns ?? 1;
			continue;
		}
		map.set(block.id, cols);
	}
	return map;
}

/** A `<!-- docx:layout pN … warn="…" -->` hazard hint for a tab-aligned line
 * whose right-hand content can wrap — a render-only break Markdown shows as a
 * clean `\t`. Two shapes, both deviation-only (fire only on the risky combo):
 *  - a tab in a MULTI-COLUMN section (the column is narrow), or
 *  - a tab right-positioned by a LEFT/center tab stop near the margin with no
 *    RIGHT tab (the résumé "San / Francisco, CA" break — short placeholders fit
 *    the gap to the margin, real values overflow and wrap; a right tab wouldn't). */
/** The consolidated top-of-read summary for every line whose tab-aligned content
 *  wraps in render, with a SINGLE command that cures them all. Empty when there's
 *  nothing wrapping. The cure spans the min..max paragraph in document order; the
 *  range form of `--tabs` only touches paragraphs that already have tab stops, so
 *  a wide range safely skips the non-tab paragraphs in between. */
function formatWrappingTabSummary(ids: string[]): string {
	if (ids.length === 0) return "";
	const count = `${ids.length} line${ids.length > 1 ? "s" : ""}`;
	const range = tabCureRange(ids);
	const fix = range
		? `edit FILE --at ${range} --tabs right`
		: ids.map((id) => `edit FILE --at ${id} --tabs right`).join(" ; ");
	return formatNote(
		"layout",
		[
			["wrap", count],
			[
				"warn",
				"tab-aligned content (e.g. dates/locations) overflows the right margin and WRAPS in render until cured",
			],
			["fix-all", fix],
		],
		ids,
	);
}

/** `pMin-pMax` (or `pMin` when one) over the simple `pN` ids in `ids`; null when
 *  none are plain body-paragraph locators (e.g. all cell-paragraph locators), in
 *  which case the caller lists per-line cures instead. */
function tabCureRange(ids: string[]): string | null {
	const nums = ids
		.map((id) => /^p(\d+)$/.exec(id)?.[1])
		.filter((value): value is string => value !== undefined)
		.map(Number);
	if (nums.length === 0) return null;
	const min = Math.min(...nums);
	const max = Math.max(...nums);
	return min === max ? `p${min}` : `p${min}-p${max}`;
}

function layoutHazardNote(paragraph: Paragraph, ctx: RenderContext): string {
	if (!paragraph.runs.some((run) => run.type === "tab")) return "";

	const cols = ctx.governingColumns.get(paragraph.id) ?? 1;
	if (cols > 1) {
		return ` ${formatNote(
			"layout",
			[
				["cols", cols],
				[
					"warn",
					`tab alignment can wrap mid-line in this ${cols}-column section; render to verify`,
				],
			],
			[paragraph.id],
		)}`;
	}

	// Fragile right-alignment via a left/center tab near the margin. List items
	// are excluded: their tab stops are the structural bullet-to-text gap, and the
	// `--tabs right` cure can't safely touch them (it'd jump the bullet text to the
	// margin), so flagging them would advertise a cure that corrupts — keep flag
	// and cure in lockstep (`scopeRangeProps` skips list paragraphs too).
	if (paragraph.list) return "";
	const tabs = paragraph.tabStops ?? [];
	if (tabs.some((tab) => tab.align === "right")) return ""; // robust
	const textWidthTwips = Math.round(ctx.contentWidthEmu / EMU_PER_TWIP);
	const fragile = tabs.find(
		(tab) =>
			(tab.align === "left" || tab.align === "center" || tab.align === "") &&
			tab.pos > textWidthTwips * 0.7,
	);
	if (!fragile) return "";
	// Record it so renderMarkdown can emit ONE consolidated top-of-doc summary with
	// a single-call cure across all wrapping lines — per-line hints alone get
	// ignored (both Haiku and Sonnet blew past them across repeated reads).
	ctx.wrappingTabLines.push(paragraph.id);
	const gapIn = twipsToInches(Math.max(0, textWidthTwips - fragile.pos));
	return ` ${formatNote(
		"layout",
		[
			["tab", `${fragile.align || "left"}@${twipsToInches(fragile.pos)}in`],
			[
				"warn",
				`right content on a LEFT tab ~${gapIn}in from the margin — wraps when longer (cure: --tabs right; see the docx:layout fix-all at top)`,
			],
		],
		[paragraph.id],
	)}`;
}

/** The locator range a section break governs. Per ECMA-376 a `<w:sectPr>`
 * applies to the content ENDING at it — everything back to the previous section
 * boundary, i.e. the paragraphs ABOVE the break, not below. That direction is
 * the off-by-one weak agents miss (they read the annotation as governing what
 * follows; the `eliot-journal` scenario put a 2-column boundary before the poems
 * and got single-column poems). Returns the `pX..pY` span (or a single id), or
 * undefined when the section governs no addressable content. */
function governedRange(blocks: Block[], index: number): string | undefined {
	let first: string | undefined;
	let last: string | undefined;
	for (let cursor = index - 1; cursor >= 0; cursor--) {
		const block = blocks[cursor];
		if (!block) continue;
		if (block.type === "sectionBreak") break;
		first = block.id;
		if (last === undefined) last = block.id;
	}
	if (!first || !last) return undefined;
	return first === last ? first : `${first}..${last}`;
}

/** A section break as an own-line `<!-- docx:section sN cols="2" type="…" -->`
 * annotation — NOT the bare `---` it used to be. `---` round-trips as a thematic
 * break (`<HorizontalRule>` → a border paragraph), so emitting it for a section
 * silently corrupted layout on `read → create` AND was indistinguishable from a
 * real thematic break. The `docx:section` comment is a read-time VISIBILITY hint:
 * the importer drops it (no reconstruction — `--ast` is the lossless view, and
 * `docx sections` / edit-in-place manage layout), and a
 * hand-authored `---` now unambiguously means a thematic break. cols/type are
 * shown for the section; document geometry rides the leading `docx:page` note. */
function renderSectionBreak(
	block: SectionBreak,
	governs?: string,
	marginalLines?: string[],
): string {
	const pairs: NotePair[] = [];
	if (block.columns !== undefined && block.columns > 1) {
		pairs.push(["cols", block.columns]);
	}
	if (block.sectionType !== undefined) pairs.push(["type", block.sectionType]);
	// EVERY section break gets a `docx:section sN` marker — including a contentless
	// one (default single-column, no explicit type; geometry rides the `docx:page`
	// note) and the trailing mandatory sectPr — so the read consistently shows
	// where each section begins (and the sN locator an agent addresses). cols/type
	// appear only when they deviate; a bare section is just `<!-- docx:section sN -->`.
	// State the scope only on a section that already deviates (multi-column or a
	// non-default type) — it's exactly the layout-bearing sections where the
	// "which side?" trap bites. The note renders at the section's START, so
	// `(below)` spells out that the columns/type govern the content that FOLLOWS.
	if (governs !== undefined && pairs.length > 0) {
		pairs.push(["applies-to", `${governs} (below)`]);
	}
	const note = formatNote("section", pairs, [block.id]);
	// Co-locate this section's per-section header/footer hints right under its
	// break, so they read next to the content they govern.
	if (marginalLines && marginalLines.length > 0) {
		return [note, ...marginalLines].join("\n");
	}
	return note;
}

/** Canonical default page geometry — US Letter portrait, 1″ margins. Geometry
 * matching this is suppressed from the `docx:page` note (deviation-only: a note
 * that repeats the default is noise). */
const DEFAULT_PAGE = {
	width: 12240,
	height: 15840,
	margin: 1440,
} as const;

/** 635 EMU per twip — the OOXML geometry conversion (1 twip = 1/20 pt). */
const EMU_PER_TWIP = 635;

/** Usable page width in EMU (page width − L/R margins) from the document
 * geometry, falling back to US-Letter-1″ (6.5″) when geometry is absent. Used to
 * flag images wider than the text column. */
function contentWidthEmu(geometry: SectionBreak | undefined): number {
	const width = geometry?.pageWidth ?? DEFAULT_PAGE.width;
	const left = geometry?.marginLeft ?? DEFAULT_PAGE.margin;
	const right = geometry?.marginRight ?? DEFAULT_PAGE.margin;
	const contentTwips = width - left - right;
	const fallback =
		(DEFAULT_PAGE.width - 2 * DEFAULT_PAGE.margin) * EMU_PER_TWIP;
	return contentTwips > 0 ? contentTwips * EMU_PER_TWIP : fallback;
}

/** The `<!-- docx:image imgN size="6.2x4.1in" … -->` annotation for an image, or
 * "" when there's nothing to declare. `size` is always shown (always
 * decision-relevant); `overflow`/`float`/`wrap`/`align` are deviation-only
 * (an inline, in-bounds image emits just its size). A DROPPED read-time hint. */
function formatImageNote(run: ImageRun, contentEmu: number): string {
	const pairs: NotePair[] = [];
	if (run.widthEmu && run.heightEmu) {
		pairs.push([
			"size",
			`${emuToInches(run.widthEmu)}x${emuToInches(run.heightEmu)}in`,
		]);
	}
	if (run.floating) pairs.push(["float", "yes"]);
	if (run.wrap) pairs.push(["wrap", run.wrap]);
	if (run.align) pairs.push(["align", run.align]);
	if (run.widthEmu && run.widthEmu > contentEmu)
		pairs.push(["overflow", "yes"]);
	if (pairs.length === 0) return "";
	return ` ${formatNote("image", pairs, [run.id])}`;
}

/** The section break carrying the document-wide geometry — the trailing
 * Returns the FIRST geometry-bearing section (page 1 — what the reader sees first
 * and the most intuitive "document geometry"), plus whether any later
 * geometry-bearing section DIFFERS from it. The `varies` flag is the honesty
 * guard: a multi-section doc can mix page setups (e.g. one landscape section), and
 * a single doc-level note would otherwise claim page 1's geometry for the whole
 * file — the exact trap that read "landscape" while page 1 rendered portrait.
 * Keying on any-geometry (not page width alone) catches a margins-only sectPr that
 * inherits `<w:pgSz>` (legal — pgSz is schema-optional). */
function documentGeometry(
	blocks: Block[],
): { geometry: SectionBreak; varies: boolean } | undefined {
	const geometric = blocks.filter(
		(block): block is SectionBreak =>
			block.type === "sectionBreak" && hasGeometry(block),
	);
	const first = geometric[0];
	if (!first) return undefined;
	const signature = geometrySignature(first);
	const varies = geometric.some(
		(block) => geometrySignature(block) !== signature,
	);
	return { geometry: first, varies };
}

/** A comparable signature of a section's RESOLVED page geometry (defaults filled,
 * size orientation-normalized) — used to detect whether sections share one page
 * setup. Mirrors the fields `formatPageNote` surfaces. */
function geometrySignature(block: SectionBreak): string {
	const width = block.pageWidth ?? DEFAULT_PAGE.width;
	const height = block.pageHeight ?? DEFAULT_PAGE.height;
	const orientation =
		block.pageOrientation ?? (width > height ? "landscape" : "portrait");
	const short = Math.min(width, height);
	const long = Math.max(width, height);
	const top = block.marginTop ?? DEFAULT_PAGE.margin;
	const right = block.marginRight ?? DEFAULT_PAGE.margin;
	const bottom = block.marginBottom ?? DEFAULT_PAGE.margin;
	const left = block.marginLeft ?? DEFAULT_PAGE.margin;
	return [orientation, short, long, top, right, bottom, left].join(",");
}

/** True when a section break declares any page geometry (size, orientation, or
 * any margin) — the discriminator for `documentGeometry`. */
function hasGeometry(block: SectionBreak): boolean {
	return (
		block.pageWidth !== undefined ||
		block.pageHeight !== undefined ||
		block.pageOrientation !== undefined ||
		block.marginTop !== undefined ||
		block.marginRight !== undefined ||
		block.marginBottom !== undefined ||
		block.marginLeft !== undefined
	);
}

/** A leading `<!-- docx:page … -->` note declaring document page geometry, but
 * ONLY the parts that deviate from the canonical default (landscape, non-Letter
 * size, non-1″ margins). Returns "" for a plain default-Letter document so the
 * common case stays clean. Includes `text-width` (the usable column width =
 * page width − left/right margins) since that's the single most decision-relevant
 * number for layout and image-overflow reasoning. A DROPPED read-time hint —
 * geometry survives edits via in-place mutation; `read --ast` carries exact
 * twips. */
function formatPageNote(
	document: { geometry: SectionBreak; varies: boolean } | undefined,
): string {
	if (!document) return "";
	const { geometry, varies } = document;
	const width = geometry.pageWidth ?? DEFAULT_PAGE.width;
	const height = geometry.pageHeight ?? DEFAULT_PAGE.height;
	const left = geometry.marginLeft ?? DEFAULT_PAGE.margin;
	const right = geometry.marginRight ?? DEFAULT_PAGE.margin;
	const top = geometry.marginTop ?? DEFAULT_PAGE.margin;
	const bottom = geometry.marginBottom ?? DEFAULT_PAGE.margin;
	const orientation =
		geometry.pageOrientation ?? (width > height ? "landscape" : "portrait");

	const pairs: NotePair[] = [];
	if (orientation === "landscape") pairs.push(["orientation", "landscape"]);
	// "Not Letter" is orientation-agnostic — compare the unordered dimension pair.
	const [short, long] = width <= height ? [width, height] : [height, width];
	if (short !== DEFAULT_PAGE.width || long !== DEFAULT_PAGE.height) {
		pairs.push(["size", `${twipsToInches(width)}x${twipsToInches(height)}in`]);
	}
	if (
		top !== DEFAULT_PAGE.margin ||
		right !== DEFAULT_PAGE.margin ||
		bottom !== DEFAULT_PAGE.margin ||
		left !== DEFAULT_PAGE.margin
	) {
		if (top === bottom && left === right && top === left) {
			pairs.push(["margins", `${twipsToInches(top)}in`]);
		} else {
			pairs.push([
				"margins",
				`${twipsToInches(top)},${twipsToInches(right)},${twipsToInches(bottom)},${twipsToInches(left)}in`,
			]);
		}
	}
	// Nothing deviated AND geometry is uniform → no note (plain default Letter).
	// When sections VARY, emit even if page 1 is default, so the note never
	// implies a uniform geometry the document doesn't have.
	if (pairs.length === 0 && !varies) return "";
	pairs.push(["text-width", `${twipsToInches(width - left - right)}in`]);
	// `varies` warns that LATER sections differ from page 1 (this note describes
	// page 1 / the leading section only) — re-apply per section, or `read --ast`.
	if (varies) pairs.push(["varies", "by-section"]);
	// Lead with the section's locator (like docx:cell/docx:p) so an agent can
	// re-apply: `docx sections --at sN --orientation/--size/--margins …`. The
	// orientation/size/margins values above are emitted in the exact flag units.
	return formatNote("page", pairs, [geometry.id]);
}

/** Build the `<!-- docx:header … -->` / `<!-- docx:footer … -->` hints, split into
 *  UNIFORM (identical across every section — head block, no section token) and
 *  PER-SECTION (text differs by section — co-located with each section break,
 *  keyed by `sN`). Deviation-only: the `type` attr is omitted when `default`.
 *  Content rides the `text` attribute so the importer's `docx:`-comment drop can't
 *  re-inject it into the body. */
function formatMarginalNotes(
	headers: Marginal[],
	footers: Marginal[],
	sectionCount: number,
): { uniform: string[]; bySection: Map<string, string[]> } {
	const uniform: string[] = [];
	const bySection = new Map<string, string[]>();
	const emit = (sectionId: string | null, line: string): void => {
		if (sectionId === null) {
			uniform.push(line);
			return;
		}
		const group = bySection.get(sectionId) ?? [];
		group.push(line);
		bySection.set(sectionId, group);
	};
	collectMarginalNotes("header", headers, sectionCount, emit);
	collectMarginalNotes("footer", footers, sectionCount, emit);
	return { uniform, bySection };
}

function collectMarginalNotes(
	noteType: "header" | "footer",
	marginals: Marginal[],
	sectionCount: number,
	emit: (sectionId: string | null, line: string) => void,
): void {
	const byType = new Map<string, Marginal[]>();
	for (const marginal of marginals) {
		const group = byType.get(marginal.type) ?? [];
		group.push(marginal);
		byType.set(marginal.type, group);
	}
	for (const [placement, group] of byType) {
		const distinctText = new Set(group.map((marginal) => marginal.text));
		const typePairs: NotePair[] =
			placement === "default" ? [] : [["type", placement]];
		// Document chrome (ONE head line, no sN) ONLY when the marginal is identical
		// AND present on EVERY section. A marginal on a SUBSET of sections — even
		// with identical text — gets per-section lines keyed by `sN`, so it's never
		// mislabeled as covering sections it doesn't (e.g. a header only on the
		// trailing section would otherwise imply the earlier sections carry it too).
		const coversEverySection = group.length >= sectionCount;
		if (distinctText.size === 1 && coversEverySection) {
			const first = group[0];
			if (!first) continue;
			emit(null, formatNote(noteType, [...typePairs, ["text", first.text]]));
			continue;
		}
		// Differs by section, or only covers some → one line per section, co-located
		// via its `sN`.
		for (const marginal of group) {
			emit(
				marginal.sectionId,
				formatNote(
					noteType,
					[...typePairs, ["text", marginal.text]],
					[marginal.sectionId],
				),
			);
		}
	}
}

function isCodeBlockParagraph(block: Block): block is Paragraph {
	return block.type === "paragraph" && isCodeBlockStyleId(block.style);
}

/** Collapse a run of CodeBlock paragraphs into a GFM fenced block. Token
 *  formatting (the colors lowlight applied on insert) gets stripped — the
 *  fenced rendering loses syntax-highlighting fidelity but stays a faithful
 *  source code representation, which is the right trade-off for a markdown
 *  view. Locator comments sit on their own lines bracketing the fence (open id
 *  before, close id after) — never on the fence lines themselves.
 *  The language tag on the opening fence comes from the first paragraph's
 *  `CodeBlock-LANG` pStyle suffix (or empty for the bare `CodeBlock`).
 *
 *  Tracked-change references inside the group's runs (someone edited a line
 *  under tracking) are collected into `ctx.referencedTrackedChanges` so the
 *  current-view footnote appendix still surfaces them — even though their
 *  CriticMarkup wrappers are stripped from the fenced rendering itself. */
function renderCodeBlockGroup(
	paragraphs: Paragraph[],
	ctx: RenderContext,
): string {
	if (ctx.options.view !== "baseline" && ctx.options.view !== "accepted") {
		// `current` view: collect tracked-change refs so [^tcN] definitions
		// still render in the footnote appendix.
		for (const paragraph of paragraphs) {
			for (const run of paragraph.runs) {
				if (run.type === "text" && run.trackedChange) {
					ctx.referencedTrackedChanges.set(
						run.trackedChange.id,
						run.trackedChange,
					);
				}
			}
		}
	}
	const lines = paragraphs.map((paragraph) =>
		paragraph.runs
			.filter(
				(run): run is TextRun =>
					run.type === "text" && typeof run.text === "string",
			)
			.map((run) => run.text)
			.join(""),
	);
	const firstId = paragraphs[0]?.id ?? "";
	const lastId = paragraphs[paragraphs.length - 1]?.id ?? firstId;
	const language = codeBlockLanguageFromStyleId(paragraphs[0]?.style) ?? "";
	// Locator comments go on their OWN lines, never on the fence lines. Gluing
	// one to the opening info string (` ```ts<!--p--> `) corrupts the language
	// id (the parser — and our own importer — reads `ts<!--` as the language);
	// putting one after the closing fence on the same line (` ``` <!--p--> `)
	// violates CommonMark (a closing fence may be followed only by spaces), so
	// the block never closes. The opening locator precedes the fence; the
	// closing locator (only when the group spans >1 source paragraph) follows it.
	const parts = [`<!-- ${firstId} -->`, `\`\`\`${language}`, ...lines, "```"];
	if (firstId !== lastId) parts.push(`<!-- ${lastId} -->`);
	return parts.join("\n");
}

function renderParagraph(
	paragraph: Paragraph,
	ctx: RenderContext,
): string | null {
	const view = ctx.options.view ?? "accepted";
	const mask = inlineEscapeMask(
		paragraphContent(paragraph.runs, view),
		hasEquationRun(paragraph.runs),
	);
	const rendered = renderRuns(paragraph.id, paragraph.runs, ctx, mask, 0);
	if (rendered.length === 0) return null;
	const prefix = paragraphPrefix(paragraph, orderedOrdinal(paragraph, ctx));
	// Trim trailing spaces/tabs (not newlines) so the single space separating
	// the body from its ` <!-- pN -->` locator doesn't accumulate one extra
	// space on every read → create → read cycle.
	const body = rendered.replace(/[ \t]+$/, "");
	// Locator + metadata. The `docx:p` note already carries the locator as its
	// leading token (like `docx:cell`), so when it's present the bare `<!-- pN -->`
	// would just duplicate `pN` — emit only the note. Plain paragraphs (no note)
	// get the bare locator. Either way every paragraph shows its addressable pN.
	const note = formatParagraphNote(paragraph);
	const locatorOrNote = note !== "" ? note : ` <!-- ${paragraph.id} -->`;
	const trailing = `${locatorOrNote}${layoutHazardNote(paragraph, ctx)}`;
	// Display equations need their trailing comments on a new line for KaTeX-based
	// renderers (Obsidian, VS Code preview, …) to recognize `$$…$$` as display
	// math — a comment after a space on the same line leaves the trailing `$`
	// looking like an unmatched math-mode toggle. Each piece begins with a leading
	// space; swap the first for the newline in that case.
	if (isDisplayEquationOnly(body)) {
		return `${prefix}${body}\n${trailing.replace(/^ /, "")}`;
	}
	return `${prefix}${body}${trailing}`;
}

/** The `<!-- docx:p pN style="Caption" align="center" -->` annotation, or "" when
 * the paragraph has nothing non-default. `style` is shown only for styles NOT
 * already conveyed by the Markdown construct (headings `#`, lists `-`, quotes
 * `>`, code fences) — so Caption / custom paragraph styles surface; `align` only
 * when it isn't the default left. A DROPPED read-time hint. */
function formatParagraphNote(paragraph: Paragraph): string {
	const pairs: NotePair[] = [];
	if (paragraph.style && !isConstructStyle(paragraph.style)) {
		pairs.push(["style", paragraph.style]);
	}
	if (paragraph.alignment && paragraph.alignment !== "left") {
		pairs.push(["align", paragraph.alignment]);
	}
	// Direct spacing/indent overrides — deviation-only by construction (the reader
	// surfaces these fields only when set directly on the paragraph, and skips a
	// list/quote's structural indent). Values use the same units as the flags so an
	// agent can read the note and re-apply: points for spacing, inches for indent,
	// a multiple for line spacing.
	const spacing = paragraph.spacing;
	if (spacing) {
		if (spacing.before !== undefined)
			pairs.push(["space-before", twipsToPoints(spacing.before)]);
		if (spacing.after !== undefined)
			pairs.push(["space-after", twipsToPoints(spacing.after)]);
		if (spacing.line !== undefined) {
			// "auto" → a bare multiple (1.5); "exact"/"atLeast" → a point value
			// tagged with the rule ("18pt exact"). Both forms feed straight back
			// through `--line-spacing` (parseLineSpacing accepts each), so the note
			// round-trips for all three rules.
			const rule = spacing.lineRule ?? "auto";
			pairs.push([
				"line-spacing",
				rule === "auto"
					? String(Number.parseFloat((spacing.line / 240).toFixed(2)))
					: `${twipsToPoints(spacing.line)} ${rule}`,
			]);
		}
	}
	const indent = paragraph.indent;
	if (indent) {
		if (indent.left !== undefined)
			pairs.push(["indent-left", `${twipsToInches(indent.left)}in`]);
		if (indent.right !== undefined)
			pairs.push(["indent-right", `${twipsToInches(indent.right)}in`]);
		if (indent.firstLine !== undefined)
			pairs.push(["first-line", `${twipsToInches(indent.firstLine)}in`]);
		if (indent.hanging !== undefined)
			pairs.push(["hanging", `${twipsToInches(indent.hanging)}in`]);
	}
	if (pairs.length === 0) return "";
	return ` ${formatNote("p", pairs, [paragraph.id])}`;
}

/** Styles already conveyed by the Markdown construct itself (so annotating them
 * would be noise): headings (`#`), code blocks (fences), and the quote/list
 * baseline styles, plus the `Normal` default. Everything else (Caption, custom
 * pStyles) is invisible in GFM and worth surfacing. */
function isConstructStyle(style: string): boolean {
	if (headingLevelFor(style) !== null) return true;
	if (isCodeBlockStyleId(style)) return true;
	return CONSTRUCT_STYLES.has(style);
}

const CONSTRUCT_STYLES: ReadonlySet<string> = new Set([
	"Normal",
	"Quote",
	"IntenseQuote",
	"QuoteListParagraph",
	"ListParagraph",
]);

/** True if the rendered body is a single display-math expression (`$$…$$`)
 *  with no other content — the case where we put the locator on its own
 *  line so KaTeX-based markdown renderers process the math correctly. */
function isDisplayEquationOnly(body: string): boolean {
	const trimmed = body.trim();
	if (!trimmed.startsWith("$$") || !trimmed.endsWith("$$")) return false;
	// Reject `$$X$$Y$$Z$$` (two separate display equations) — only single ones.
	const inner = trimmed.slice(2, -2);
	return !inner.includes("$$");
}

/** The 1-based ordinal for an ordered-list paragraph, counting items per
 * `${numId}:${level}` across the document (Word numbers a list continuously by
 * numId). Bumping a level resets all deeper levels so a nested sub-list restarts.
 * Returns 1 for non-ordered paragraphs (unused by the caller). */
function orderedOrdinal(paragraph: Paragraph, ctx: RenderContext): number {
	if (!paragraph.list?.ordered) return 1;
	const { numId, level } = paragraph.list;
	const key = `${numId}:${level}`;
	const next = (ctx.orderedCounters.get(key) ?? 0) + 1;
	ctx.orderedCounters.set(key, next);
	for (const existing of ctx.orderedCounters.keys()) {
		const [keyNumId, keyLevel] = existing.split(":");
		if (keyNumId === String(numId) && Number(keyLevel) > level) {
			ctx.orderedCounters.delete(existing);
		}
	}
	return next;
}

function paragraphPrefix(paragraph: Paragraph, ordinal: number): string {
	// Blockquote prefix comes before everything else. `> ` repeats per
	// nesting depth; the AST reader fills `quoteDepth` from `pStyle="Quote"`
	// / `pStyle="QuoteListParagraph"` plus the paragraph's `<w:ind w:left>`
	// value. Markdown stitches adjacent `> ` lines back into one logical
	// blockquote on re-parse.
	const quotePrefix = paragraph.quoteDepth
		? "> ".repeat(paragraph.quoteDepth)
		: "";
	const headingLevel = headingLevelFor(paragraph.style);
	if (headingLevel !== null) {
		return `${quotePrefix}${"#".repeat(headingLevel)} `;
	}
	if (paragraph.list) {
		const indent = "  ".repeat(paragraph.list.level);
		// Emit the real ordinal (1. 2. 3.) so the RAW markdown reads correctly —
		// a uniform `1.` only auto-increments once a renderer displays it, leaving
		// the text an agent/human reads as a confusing wall of `1.`.
		const marker = paragraph.list.ordered ? `${ordinal}. ` : "- ";
		const task =
			paragraph.taskState === "checked"
				? "[x] "
				: paragraph.taskState === "unchecked"
					? "[ ] "
					: "";
		return `${quotePrefix}${indent}${marker}${task}`;
	}
	return quotePrefix;
}

function headingLevelFor(style: string | undefined): number | null {
	if (!style) return null;
	if (!style.startsWith("Heading")) return null;
	const remainder = style.slice("Heading".length).trim();
	if (remainder === "") return 1;
	const parsed = Number(remainder);
	if (!Number.isInteger(parsed) || parsed < 1 || parsed > 6) return null;
	return parsed;
}

function renderRuns(
	paragraphId: string,
	runs: Run[],
	ctx: RenderContext,
	mask: boolean[],
	baseOffset: number,
): string {
	const view = ctx.options.view ?? "accepted";
	const visibleEntries: { run: Run; originalIndex: number }[] = [];
	runs.forEach((run, index) => {
		if (run.type === "text" && !isRunVisible(run, view)) return;
		visibleEntries.push({ run, originalIndex: index });
	});

	let out = "";
	let cursor = 0;
	// Where the next non-code text run sits in the escape `mask` (which spans the
	// whole pairing scope — the paragraph, or for a cell, every paragraph in it).
	// Code runs and non-text runs aren't in the parsed content, so they don't
	// advance it — see `paragraphContent`.
	let contentCursor = baseOffset;
	while (cursor < visibleEntries.length) {
		const entry = visibleEntries[cursor];
		if (!entry) {
			cursor++;
			continue;
		}
		const { run } = entry;
		if (run.type === "text") {
			let lookahead = cursor + 1;
			while (lookahead < visibleEntries.length) {
				const next = visibleEntries[lookahead];
				if (!next || next.run.type !== "text") break;
				if (!sameDecoration(run, next.run)) break;
				lookahead++;
			}
			const segment = visibleEntries.slice(cursor, lookahead);
			const segmentRuns = segment.map((entry) => entry.run as TextRun);
			out += renderTextSegment(
				segmentRuns,
				view,
				ctx.baseline,
				mask,
				contentCursor,
			);
			// Inline code is excluded from the parsed content (literal in backticks),
			// so only non-code segments advance the mask cursor.
			if (segmentRuns[0]?.runStyle !== "Code") {
				contentCursor += segmentRuns.reduce(
					(sum, segmentRun) => sum + segmentRun.text.length,
					0,
				);
			}
			if (view === "current") {
				for (const segmentRun of segmentRuns) {
					if (segmentRun.trackedChange) {
						ctx.referencedTrackedChanges.set(
							segmentRun.trackedChange.id,
							segmentRun.trackedChange,
						);
					}
				}
			}
			out += commentEndingsFor(paragraphId, segment, ctx.commentIndex);
			cursor = lookahead;
			continue;
		}
		if (run.type === "image") {
			const alt = sanitizeAltText(run.alt ?? run.id);
			// Content-addressed URL: `<sha256>.<ext>`. The walker on the
			// import side (`@core/markdown::preloadImages`) recognizes the
			// shape and reuses the existing media part by hash instead of
			// shelling out to `loadImageSource` — that's what makes
			// `read → edit → write` round-trip without re-fetching images.
			// `imageById.values()` in `Body` shares the same hash → rId
			// mapping the walker queries. Mirrors the on-disk naming
			// convention used by `docx images extract`.
			const extension = extensionForImageMime(run.contentType) ?? "bin";
			out += `![${alt}](${run.hash}.${extension})`;
			// Size (always) + float/wrap/align/overflow (deviation-only): the
			// `![](hash)` proves existence, not "6.2in wide, past the margin".
			out += formatImageNote(run, ctx.contentWidthEmu);
		} else if (run.type === "break") {
			// A line break renders as a real newline so verse/line-structured
			// text round-trips (import maps a soft newline back to <w:br/>). In
			// table cells, escapeCell converts these newlines to <br> so the
			// markdown table stays one row per record.
			if (run.kind === "line") out += "\n";
		} else if (run.type === "tab") {
			out += "\t";
		} else if (run.type === "equation") {
			// `$…$` for inline, `$$…$$` for display. The walker in
			// `@core/equation` reconstructed `run.latex` from the OMML subtree;
			// `run.text` is the legacy plaintext fallback for fully-degraded
			// (unrecognized) equations — use it only when the LaTeX walker
			// returned empty.
			const body = run.latex.length > 0 ? run.latex : run.text;
			out += run.display ? `$$${body}$$` : `$${body}$`;
			// Append any comment endings that close on this equation run
			// (audit comments from tracked equation edits anchor here).
			out += commentEndingsFor(
				paragraphId,
				[{ originalIndex: cursor }],
				ctx.commentIndex,
			);
		} else if (run.type === "noteRef") {
			if (run.kind === "footnote") ctx.referencedFootnoteIds.add(run.id);
			else ctx.referencedEndnoteIds.add(run.id);
			out += `[^${run.id}]`;
		} else if (run.type === "chart") {
			out += `\`[${run.kind}]\``;
		}
		cursor++;
	}
	return out;
}

function commentEndingsFor(
	paragraphId: string,
	segment: { originalIndex: number }[],
	commentIndex: CommentIndex,
): string {
	if (commentIndex.endingsByRun.size === 0) return "";
	let out = "";
	for (const entry of segment) {
		const ids = commentIndex.endingsByRun.get(
			slotKey(paragraphId, entry.originalIndex),
		);
		if (!ids) continue;
		for (const commentId of ids) out += `[^${commentId}]`;
	}
	return out;
}

function sameDecoration(a: TextRun, b: TextRun): boolean {
	return (
		(a.bold ?? false) === (b.bold ?? false) &&
		(a.italic ?? false) === (b.italic ?? false) &&
		(a.strike ?? false) === (b.strike ?? false) &&
		(a.underline ?? "") === (b.underline ?? "") &&
		(a.underlineColor ?? "") === (b.underlineColor ?? "") &&
		(a.color ?? "") === (b.color ?? "") &&
		(a.colorTheme ?? "") === (b.colorTheme ?? "") &&
		(a.colorThemeTint ?? "") === (b.colorThemeTint ?? "") &&
		(a.colorThemeShade ?? "") === (b.colorThemeShade ?? "") &&
		(a.highlight ?? "") === (b.highlight ?? "") &&
		(a.shade ?? "") === (b.shade ?? "") &&
		(a.font ?? "") === (b.font ?? "") &&
		(a.sizeHalfPoints ?? 0) === (b.sizeHalfPoints ?? 0) &&
		(a.vertAlign ?? "") === (b.vertAlign ?? "") &&
		(a.smallCaps ?? false) === (b.smallCaps ?? false) &&
		(a.allCaps ?? false) === (b.allCaps ?? false) &&
		(a.runStyle ?? "") === (b.runStyle ?? "") &&
		a.hyperlink?.id === b.hyperlink?.id &&
		a.trackedChange?.id === b.trackedChange?.id &&
		sameCommentSet(a.comments, b.comments)
	);
}

function sameCommentSet(
	left: string[] | undefined,
	right: string[] | undefined,
): boolean {
	const a = left ?? [];
	const b = right ?? [];
	if (a.length !== b.length) return false;
	for (let i = 0; i < a.length; i++) {
		if (a[i] !== b[i]) return false;
	}
	return true;
}

function renderTextSegment(
	runs: TextRun[],
	view: MarkdownView,
	baseline: RunFormatBaseline,
	mask: boolean[],
	offset: number,
): string {
	const text = runs.map((run) => run.text).join("");
	if (text.length === 0) return "";
	const first = runs[0];
	if (!first) return "";
	// Markdown emphasis around whitespace-only text (`** **`) mis-parses — a
	// reader pairs the two `**` and bolds everything between them. So a blank run
	// carries its emphasis as unambiguous HTML tags (`<b>`/`<i>`/`<s>`) instead,
	// which both renders correctly AND round-trips (consistent with the
	// underline/highlight/color a blank run already keeps via the wrapper below).
	const isBlank = text.trim().length === 0;
	const wrap = needsHtmlWrap(first, baseline) || isBlank;
	// Escape only the characters the importer's parser would CONSUME as markup —
	// `mask` is the per-character verdict from `inlineEscapeMask`, so a paired `$`
	// or a link-forming `[` is escaped while an inert `[ x ]` checkbox or a lone
	// `$5` stays byte-clean. remark decodes the escapes on import, so the text
	// round-trips exact. Inline code is the exception: markdown isn't parsed
	// between backticks, so its text rides verbatim (and was left out of `mask`).
	const isCode = first.runStyle === "Code";
	let out = isCode ? text : applyEscapeMask(text, mask, offset);
	// Backticks INSIDE other formatting per GFM precedence — `**`x`**` is bold
	// code; `**x**` inside backticks would be literal asterisks. Skip Code runs
	// inside a fenced code block (callers strip runStyle on those — see
	// `renderCodeBlockGroup`); this branch handles `runStyle: "Code"` only when
	// it's still meaningful (inline code spans in a normal paragraph). Code spans
	// on whitespace are valid (backticks have no flanking rule), so no split.
	if (isCode) out = `\`${out}\``;
	if (isBlank) {
		if (first.bold) out = `<b>${out}</b>`;
		if (first.italic) out = `<i>${out}</i>`;
		if (first.strike) out = `<s>${out}</s>`;
	} else {
		if (first.bold) out = `**${out}**`;
		if (first.italic) out = `*${out}*`;
		if (first.strike) out = `~~${out}~~`;
	}
	// Everything else (color, highlight, shading, underline, super/sub, caps,
	// font, size, theme color) rides in HTML so a real markdown reader renders
	// it: semantic tags where they exist (<mark>/<sup>/<sub>/<u>), a `<span
	// style>` for the CSS-expressible props, and `data-*` attributes for the
	// OOXML-only bits CSS can't say. The import side parses these back;
	// `read --ast` is the lossless format. See `wrapRunFormatting`.
	if (wrap) out = wrapRunFormatting(out, first, baseline);
	if (first.hyperlink) {
		const target = first.hyperlink.url ?? `#${first.hyperlink.anchor ?? ""}`;
		out = `[${out}](${target})`;
	}
	if (view === "current" && first.trackedChange) {
		const marker = criticMarkerFor(first.trackedChange.kind);
		out = `{${marker}${out}${marker}}[^${first.trackedChange.id}]`;
	}
	return out;
}

/** CriticMarkup doesn't have a native "moved" marker, so we render moveTo
 * with the same `++` markers as an insertion (the text appears at this
 * location in the accepted view) and moveFrom with the same `--` as a
 * deletion (the text leaves this location). The footnote definition
 * carries the precise kind so a reader can distinguish move vs. ins/del. */
function criticMarkerFor(kind: TrackedChange["kind"]): "++" | "--" {
	if (kind === "ins" || kind === "moveTo") return "++";
	return "--";
}

/** Whether a run carries formatting beyond native markdown (bold/italic/strike/
 * code/link) — i.e. whether it needs an HTML wrapper. When false the run stays
 * plain markdown and its text isn't HTML-escaped. */
function needsHtmlWrap(run: TextRun, baseline: RunFormatBaseline): boolean {
	return Boolean(
		(run.color && !isDefaultColor(run.color)) ||
			(run.colorTheme && !isDefaultThemeColor(run)) ||
			run.shade ||
			(run.font && run.font !== baseline.font) ||
			(run.sizeHalfPoints !== undefined &&
				run.sizeHalfPoints !== baseline.sizeHalfPoints) ||
			run.smallCaps ||
			run.allCaps ||
			run.underline ||
			run.vertAlign === "superscript" ||
			run.vertAlign === "subscript" ||
			run.highlight,
	);
}

/** Wrap already-markdown-formatted text in the HTML that carries a run's
 * non-native formatting, innermost → outermost: `<span style + data-*>` (color,
 * font, size, caps, shade, theme color) → `<u>` (underline) → `<sup>`/`<sub>`
 * (vertical align) → `<mark>` (highlight). The nesting order is fixed so the
 * importer (`gatherHtmlSpans` in `core/markdown/inline-surgery.ts`) reverses it
 * deterministically — this pairing is the read↔import contract. CSS-expressible
 * props go in `style`; OOXML-only ones (theme color, underline style) ride as
 * `data-*` attributes a renderer ignores but the importer reads. */
function wrapRunFormatting(
	body: string,
	run: TextRun,
	baseline: RunFormatBaseline,
): string {
	const styles: string[] = [];
	const attrs: string[] = [];
	// Black / "auto" is the universal default — emitting it says nothing.
	if (run.color && !isDefaultColor(run.color))
		styles.push(`color:#${run.color}`);
	if (run.shade) styles.push(`background-color:#${run.shade}`);
	// Font and size matching the document baseline are declared once in the
	// `<!-- docx:base … -->` note and omitted here (round-trip safe — the importer
	// re-applies them); a run that deviates keeps its value.
	if (run.font && run.font !== baseline.font) {
		styles.push(`font-family:${cssFontFamily(run.font)}`);
	}
	if (
		run.sizeHalfPoints !== undefined &&
		run.sizeHalfPoints !== baseline.sizeHalfPoints
	) {
		styles.push(`font-size:${run.sizeHalfPoints / 2}pt`);
	}
	if (run.smallCaps) styles.push("font-variant:small-caps");
	if (run.allCaps) styles.push("text-transform:uppercase");
	if (run.colorTheme && !isDefaultThemeColor(run)) {
		attrs.push(htmlAttr("data-color-theme", run.colorTheme));
		if (run.colorThemeTint) {
			attrs.push(htmlAttr("data-color-theme-tint", run.colorThemeTint));
		}
		if (run.colorThemeShade) {
			attrs.push(htmlAttr("data-color-theme-shade", run.colorThemeShade));
		}
	}
	let out = body;
	if (styles.length > 0 || attrs.length > 0) {
		// `style` and `data-*` values are HTML-attribute-escaped so a crafted font
		// name (e.g. one containing `"`) can't close the attribute early and inject
		// a sibling attribute — the importer decodes the entity back verbatim.
		const stylePart =
			styles.length > 0 ? ` ${htmlAttr("style", styles.join(";"))}` : "";
		const attrPart = attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
		out = `<span${stylePart}${attrPart}>${out}</span>`;
	}
	if (run.underline === "single") {
		out = `<u>${out}</u>`;
	} else if (run.underline) {
		const color = run.underlineColor
			? ` ${htmlAttr("data-underline-color", run.underlineColor)}`
			: "";
		out = `<u ${htmlAttr("data-underline", run.underline)}${color}>${out}</u>`;
	}
	if (run.vertAlign === "superscript") out = `<sup>${out}</sup>`;
	else if (run.vertAlign === "subscript") out = `<sub>${out}</sub>`;
	if (run.highlight) {
		// `<mark>` defaults to yellow; carry any other named highlight in a
		// data attribute (the importer reads it back to the exact OOXML name).
		const named =
			run.highlight === "yellow"
				? ""
				: ` ${htmlAttr("data-highlight", run.highlight)}`;
		out = `<mark${named}>${out}</mark>`;
	}
	return out;
}

/** Black (`000000`) or `auto` — the universal default text color, which carries
 * no information when stamped explicitly on a run. */
function isDefaultColor(color: string): boolean {
	const normalized = color.toLowerCase();
	return normalized === "000000" || normalized === "auto";
}

/** `text1`/`dark1` with no tint/shade is the default dark-text theme slot
 * (black) — the theme-color equivalent of an explicit `000000`. */
function isDefaultThemeColor(run: TextRun): boolean {
	return (
		(run.colorTheme === "text1" || run.colorTheme === "dark1") &&
		!run.colorThemeTint &&
		!run.colorThemeShade
	);
}

/** A CSS `font-family` value — quote names containing whitespace so the
 * declaration stays valid (`font-family:'Times New Roman'`). */
function cssFontFamily(font: string): string {
	return /\s/.test(font) ? `'${font}'` : font;
}

/** Backslash-escape `text` per a precomputed `mask` (from `inlineEscapeMask`) —
 * `mask[offset + i]` decides character `i`. `offset` is where this run's text
 * sits in the scope-wide content the mask was built over (a paragraph, or a whole
 * table cell), so a pair straddling a `<mark>` is escaped consistently. */
function applyEscapeMask(
	text: string,
	mask: boolean[],
	offset: number,
): string {
	let out = "";
	for (let index = 0; index < text.length; index++) {
		const char = text[index];
		if (char === undefined) continue;
		out += mask[offset + index] ? `\\${char}` : char;
	}
	return out;
}

/** The text the escape mask is parsed over: every VISIBLE, non-code text run's
 * text, concatenated in order. Inline code is excluded (literal between backticks,
 * never escaped), as are non-text runs (images/equations/breaks aren't markdown
 * content) — `renderRuns` advances the mask cursor by exactly this, so the offsets
 * line up. A `<w:del>` hidden in the accepted view emits nothing, so it's skipped
 * too: it can't pair with anything. */
function paragraphContent(runs: Run[], view: MarkdownView): string {
	let content = "";
	for (const run of runs) {
		if (run.type !== "text") continue;
		if (run.runStyle === "Code") continue;
		if (!isRunVisible(run, view)) continue;
		content += run.text;
	}
	return content;
}

function hasEquationRun(runs: Run[]): boolean {
	return runs.some((run) => run.type === "equation");
}

function renderTable(table: Table, ctx: RenderContext): string | null {
	// Whole-row tracked changes are filtered by view: accepted drops deleted
	// rows, baseline drops inserted ones. Cell-level (column) tracked changes
	// can't be represented in a GFM table, so they're left in place.
	const view = ctx.options.view ?? "accepted";
	const rows = table.rows.filter((row) => isRowVisible(row, view));
	if (rows.length === 0) return null;
	const colCount = Math.max(...rows.map((row) => row.cells.length));
	if (colCount === 0) return null;
	const renderedRows = rows.map((row) => {
		const cells: string[] = [];
		for (let columnIndex = 0; columnIndex < colCount; columnIndex++) {
			const cell = row.cells[columnIndex];
			cells.push(cell ? renderCell(cell, ctx) : "");
		}
		return cells;
	});
	const lines: string[] = [];
	const headerRow = renderedRows[0];
	if (!headerRow) return null;
	lines.push(rowToLine(headerRow));
	lines.push(`| ${Array(colCount).fill("---").join(" | ")} |`);
	for (let rowIndex = 1; rowIndex < renderedRows.length; rowIndex++) {
		const row = renderedRows[rowIndex];
		if (row) lines.push(rowToLine(row));
	}
	const body = lines.join("\n");
	// A leading own-line note carries the visual structure GFM can't: uneven
	// column widths and table borders (deviation-only — even/borderless tables
	// stay clean). A DROPPED read-time hint; the importer ignores it.
	const note = formatTableNote(table);
	return note ? `${note}\n\n${body}` : body;
}

/** The `<!-- docx:table tN widths="…" borders="…" -->` note, or "" when the
 * table has nothing non-default to declare (even columns, no explicit borders). */
function formatTableNote(table: Table): string {
	const pairs: NotePair[] = [];
	const widths = unevenWidths(table.grid);
	if (widths) pairs.push(["widths", widths]);
	// `single` is the universal default (Word's default table border AND the
	// border this tool emits on every created table), so it's noise. Surface only
	// the actionable deviations: `none` (borderless), `double`, `mixed`, … .
	if (table.borders && table.borders !== "single") {
		pairs.push(["borders", table.borders]);
	}
	if (pairs.length === 0) return "";
	return formatNote("table", pairs, [table.id]);
}

/** Comma-joined column widths in inches, but ONLY when the columns are
 * meaningfully uneven (>5% spread). The GFM default is even columns, so even
 * widths carry no signal — deviation-only. Returns undefined otherwise. */
function unevenWidths(grid: number[]): string | undefined {
	if (grid.length < 2) return undefined;
	const min = Math.min(...grid);
	const max = Math.max(...grid);
	// Suppress only genuinely even grids; an all-zero/degenerate grid (max 0) is
	// caught by the spread check below. (A single 0-width column is the MOST
	// uneven case — it must surface, not be suppressed.)
	if (max - min <= max * 0.05) return undefined;
	return `${grid.map((width) => twipsToInches(width)).join(",")}in`;
}

function isRowVisible(row: TableRow, view: MarkdownView): boolean {
	const kind = row.trackedChange?.kind;
	if (view === "accepted" && kind === "rowDel") return false;
	if (view === "baseline" && kind === "rowIns") return false;
	return true;
}

function rowToLine(cells: string[]): string {
	return `| ${cells.join(" | ")} |`;
}

function renderCell(cell: TableCell, ctx: RenderContext): string {
	const parts: string[] = [];
	// A cell renders as ONE line, its paragraphs joined by `<br>`, so a pair can
	// straddle paragraphs — a `$` alone in its own paragraph partnered across the
	// `<br>`. Build the escape mask over the WHOLE cell's content with the
	// paragraphs joined by `\n` (mimicking the `<br>` line boundary): math still
	// pairs across a soft break (so the cross-paragraph `$` is escaped) but a
	// link/reference can't form across it (so a `]` ending one paragraph won't
	// fuse with a `(`/`[` opening the next into a phantom link). Each paragraph
	// gets its base offset into that content — `+ 1` per paragraph for the `\n`.
	const view = ctx.options.view ?? "accepted";
	const cellParagraphs = cell.blocks.filter(
		(block): block is Paragraph => block.type === "paragraph",
	);
	const cellContent = cellParagraphs
		.map((paragraph) => paragraphContent(paragraph.runs, view))
		.join("\n");
	const hasEquation = cellParagraphs.some((paragraph) =>
		hasEquationRun(paragraph.runs),
	);
	const mask = inlineEscapeMask(cellContent, hasEquation);
	let baseOffset = 0;
	for (const block of cell.blocks) {
		if (block.type === "paragraph") {
			const rendered = renderRuns(block.id, block.runs, ctx, mask, baseOffset);
			baseOffset += paragraphContent(block.runs, view).length + 1;
			if (rendered.length === 0) continue;
			// Trim trailing spaces/tabs so the locator separator doesn't grow
			// each round-trip (newlines become <br> via escapeCell).
			parts.push(`${rendered.replace(/[ \t]+$/, "")} <!-- ${block.id} -->`);
			continue;
		}
		if (block.type === "table") {
			parts.push(renderNestedTable(block, ctx));
		}
	}
	const body = parts.join("<br>");
	// Cell-level visual structure GFM collapses: merge (gridSpan/vMerge) and
	// background shading. Per the naming rule (bare = locator, docx: = metadata)
	// this is its own `docx:cell` annotation, NOT riding the bare cell locator. A
	// DROPPED read-time hint. A merged/shaded cell is often EMPTY (a
	// vMerge="continue" cell carries no content) — surface the note regardless.
	const note = cellNote(cell);
	if (!note) return escapeCell(body);
	return escapeCell(body ? `${body} ${note}` : note);
}

/** The `<!-- docx:cell t0:r0c0 gridSpan="2" vMerge="continue" shading="FFE699"
 * -->` annotation for a cell with merge/shading, or "" for a plain cell. Carries
 * the cell's address (the first paragraph's locator minus its `:pN`). */
function cellNote(cell: TableCell): string {
	const pairs: NotePair[] = [];
	if (cell.gridSpan && cell.gridSpan > 1)
		pairs.push(["gridSpan", cell.gridSpan]);
	if (cell.vMerge) pairs.push(["vMerge", cell.vMerge]);
	if (cell.shading) pairs.push(["shading", cell.shading]);
	if (pairs.length === 0) return "";
	const address = cellAddress(cell);
	return formatNote("cell", pairs, address ? [address] : []);
}

/** A cell's locator address (`t0:r0c0`), derived from its first block's locator
 * by dropping the trailing positional segment. A cell's first block is a
 * paragraph (`t0:r0c0:p0` → `t0:r0c0`) or a nested table (`t0:r0c0:t0` →
 * `t0:r0c0`) — keying on paragraphs alone would drop the address of a cell whose
 * only content is a nested table, leaving the `docx:cell` note unaddressable. */
function cellAddress(cell: TableCell): string | undefined {
	return cell.blocks[0]?.id.replace(/:(?:p|t)\d+$/, "");
}

function renderNestedTable(table: Table, ctx: RenderContext): string {
	const rows = table.rows.map((row) =>
		row.cells.map((cell) => renderCell(cell, ctx)).join(" / "),
	);
	return rows.join(" // ");
}

function escapeCell(text: string): string {
	// Table cells are single-line in GFM: a real newline (from a <w:br/> run or
	// a multi-paragraph cell) must become <br> so the row stays intact.
	return text.replace(/\|/g, "\\|").replace(/\n/g, "<br>");
}

function sanitizeAltText(text: string): string {
	return text.replace(/[\r\n]+/g, " ").replace(/\]/g, "\\]");
}

function renderCommentFootnotes(
	commentIndex: CommentIndex,
	comments: Comment[],
): string {
	if (commentIndex.orderedIds.length === 0) return "";
	const byId = new Map(comments.map((comment) => [comment.id, comment]));
	const sorted = [...commentIndex.orderedIds].sort(commentIdCompare);
	const lines: string[] = [];
	for (const commentId of sorted) {
		const comment = byId.get(commentId);
		if (!comment) continue;
		const span = commentIndex.spanText.get(commentId) ?? "";
		lines.push(formatFootnote(comment, span));
	}
	return lines.join("\n");
}

function renderNoteDefinitions(
	notes: Footnote[],
	referenced: Set<string>,
): string {
	if (referenced.size === 0) return "";
	const byId = new Map(notes.map((note) => [note.id, note]));
	const sorted = [...referenced].sort(noteIdCompare);
	const lines: string[] = [];
	for (const id of sorted) {
		const note = byId.get(id);
		const body = (note?.text ?? "").replace(/\s+/g, " ").trim();
		lines.push(`[^${id}]: ${body}`);
	}
	return lines.join("\n");
}

function renderTrackedChangeFootnotes(
	referenced: Map<string, TrackedChange>,
): string {
	if (referenced.size === 0) return "";
	const sorted = [...referenced.values()].sort((a, b) =>
		trackedChangeIdCompare(a.id, b.id),
	);
	const lines: string[] = [];
	for (const change of sorted) {
		const kind = trackedChangeLabelFor(change.kind);
		const author = change.author || "unknown";
		const meta = change.date ? `${author} (${change.date})` : author;
		lines.push(`[^${change.id}]: ${kind} by ${meta}`);
	}
	return lines.join("\n");
}

function trackedChangeLabelFor(kind: TrackedChange["kind"]): string {
	if (kind === "ins") return "insertion";
	if (kind === "del") return "deletion";
	if (kind === "moveTo") return "moveTo";
	return "moveFrom";
}

function trackedChangeIdCompare(left: string, right: string): number {
	return numericIdCompare(left, right, /^tc(\d+)$/);
}

function noteIdCompare(left: string, right: string): number {
	return numericIdCompare(left, right, /(\d+)$/);
}

function commentIdCompare(left: string, right: string): number {
	return numericIdCompare(left, right, /^c(\d+)$/);
}

function numericIdCompare(
	left: string,
	right: string,
	pattern: RegExp,
): number {
	const leftMatch = left.match(pattern);
	const rightMatch = right.match(pattern);
	if (leftMatch?.[1] && rightMatch?.[1]) {
		return Number(leftMatch[1]) - Number(rightMatch[1]);
	}
	return left.localeCompare(right);
}

function formatFootnote(comment: Comment, spanText: string): string {
	const quoted = quoteSpan(spanText);
	const reply = comment.parentId ? ` ↳ ${comment.parentId}` : "";
	const resolved = comment.resolved ? "✓ " : "";
	const body = comment.text.replace(/\s+/g, " ").trim();
	return `[^${comment.id}]: ${quoted} — ${resolved}${comment.author} (${comment.date})${reply}: ${body}`;
}

function quoteSpan(text: string): string {
	const collapsed = text.replace(/\s+/g, " ").trim();
	const escaped = collapsed.replace(/"/g, '\\"');
	return `"${escaped}"`;
}

function sliceBlocks(
	blocks: Block[],
	from: string | undefined,
	to: string | undefined,
): Block[] {
	if (!from && !to) return blocks;
	const fromId = from ? blockIdForLocator(from, "from") : null;
	const toId = to ? blockIdForLocator(to, "to") : null;
	const fromIndex = fromId ? blocks.findIndex((b) => b.id === fromId) : 0;
	if (from && fromId && fromIndex === -1) {
		throw new MarkdownLocatorError(
			from,
			`--from ${from} not found at document top level`,
		);
	}
	const toIndex = toId
		? blocks.findIndex((b) => b.id === toId)
		: blocks.length - 1;
	if (to && toId && toIndex === -1) {
		throw new MarkdownLocatorError(
			to,
			`--to ${to} not found at document top level`,
		);
	}
	if (toIndex < fromIndex) return [];
	return blocks.slice(fromIndex, toIndex + 1);
}

function blockIdForLocator(input: string, position: "from" | "to"): string {
	let parsed: Locator;
	try {
		parsed = parseLocator(input);
	} catch (err) {
		if (err instanceof LocatorParseError) {
			throw new MarkdownLocatorError(input, err.message);
		}
		throw err;
	}
	switch (parsed.kind) {
		case "block":
			return parsed.blockId;
		case "blockSpan":
			return parsed.blockId;
		case "range":
			return position === "from" ? parsed.start.blockId : parsed.end.blockId;
		case "blockRange":
			return position === "from" ? parsed.startBlockId : parsed.endBlockId;
		case "cell":
			return parsed.tableId;
		case "comment":
		case "image":
		case "hyperlink":
		case "trackedChange":
		case "equation":
		case "footnote":
		case "endnote":
		case "tableRow":
		case "tableColumn":
		case "cellRange":
			throw new MarkdownLocatorError(
				input,
				`--${position} does not accept a ${parsed.kind} locator — use a paragraph or table locator`,
			);
	}
}
