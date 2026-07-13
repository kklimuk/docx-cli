import { type Body, iterateBlocks } from "../ast/document/body";
import type { Block, Paragraph, TrackedChange } from "../ast/types";
import { replaceAcrossParagraphs } from "./replace-across";

export {
	replaceSpanInParagraph,
	type Span,
	type TrackedReplaceOptions,
} from "./replace-span";
export { replaceAcrossParagraphs };

export type TextMatch = {
	blockId: string;
	start: number;
	end: number;
	text: string;
	trackedChanges?: TrackedChange[];
};

export type FindView = "accepted" | "current" | "baseline";

export type FindOptions = {
	regex?: boolean;
	ignoreCase?: boolean;
	view?: FindView;
	exact?: boolean;
};

export type NormalizationKind =
	| "strip-md-emphasis"
	| "smart-quotes"
	| "dashes"
	| "whitespace"
	| "bullets";

export type FindResult = {
	matches: TextMatch[];
	normalizedQuery?: string;
	normalizationApplied?: NormalizationKind[];
};

export function findTextSpans(
	doc: Body,
	query: string,
	options: FindOptions = {},
): FindResult {
	const view = options.view ?? "accepted";
	const { matcher, ...normalization } = buildMatcher(query, options);
	const out: TextMatch[] = [];
	collectMatches(doc.blocks, matcher, view, out);
	return { matches: out, ...normalization };
}

/** How a query becomes a matcher — shared by `findTextSpans` and
 *  `findAcrossParagraphs` so the two search paths can't drift on
 *  normalization, regex, or case handling. Reports the normalized query only
 *  when normalization actually changed it, matching the `FindResult` fields. */
function buildMatcher(
	query: string,
	options: FindOptions,
): {
	matcher: Matcher;
	normalizedQuery?: string;
	normalizationApplied?: NormalizationKind[];
} {
	const ignoreCase = options.ignoreCase ?? false;
	if (options.regex) return { matcher: regexMatcher(query, ignoreCase) };
	if (options.exact) {
		return { matcher: literalMatcher(query, ignoreCase, false) };
	}
	const { normalized, applied } = normalizeQuery(query);
	const matcher = literalMatcher(normalized, ignoreCase, true);
	if (applied.length === 0) return { matcher };
	return {
		matcher,
		normalizedQuery: normalized,
		normalizationApplied: applied,
	};
}

/** A match that may span consecutive body paragraphs. `startBlockId` ===
 *  `endBlockId` for a match that stayed inside one paragraph (its "\n" hit an
 *  in-paragraph `<w:br/>`); they differ when the match crossed a paragraph
 *  boundary. Offsets are paragraph-local, in the same view coordinate space as
 *  `findTextSpans`. */
export type ParagraphSpanMatch = {
	startBlockId: string;
	startOffset: number;
	endBlockId: string;
	endOffset: number;
	text: string;
};

/** What `findAcrossParagraphs` returns: the spanning matches plus the same
 *  normalization bookkeeping `findTextSpans` reports, so both search paths
 *  surface `normalizedQuery`/`normalizationApplied` identically. */
export type FindAcrossResult = {
	matches: ParagraphSpanMatch[];
	normalizedQuery?: string;
	normalizationApplied?: NormalizationKind[];
};

/** Match `query` across CONSECUTIVE paragraphs — the path the CLI takes when
 *  the pattern contains a newline. The search space joins each maximal run of
 *  adjacent paragraph blocks with a single "\n" per boundary, the same
 *  character an in-paragraph `<w:br/>` contributes (see
 *  `paragraphTextForView`), so an agent writes one "\n" to cross one line
 *  whether it's a soft break or a paragraph mark. A table or section break
 *  ends a run — no match spans them — and each table CELL's paragraphs form
 *  their own runs, so a match can span lines inside one cell but never cross a
 *  cell wall. Normalization, regex, and view selection behave exactly as in
 *  `findTextSpans`. */
export function findAcrossParagraphs(
	doc: Body,
	query: string,
	options: FindOptions = {},
): FindAcrossResult {
	const view = options.view ?? "accepted";
	const { matcher, ...normalization } = buildMatcher(query, options);

	const out: ParagraphSpanMatch[] = [];
	for (const group of consecutiveParagraphGroups(doc.blocks)) {
		const entries: { blockId: string; flatStart: number; length: number }[] =
			[];
		let flat = "";
		for (const paragraph of group) {
			const text = paragraphTextForView(paragraph, view);
			entries.push({
				blockId: paragraph.id,
				flatStart: flat.length,
				length: text.length,
			});
			// One "\n" per boundary; the trailing one (after the last paragraph) is
			// sliced off so it never participates in a match.
			flat += `${text}\n`;
		}
		flat = flat.slice(0, -1);

		// Matcher spans arrive in ascending, non-overlapping order, so each
		// offset lookup resumes at the previous one's entry instead of rescanning
		// the group's entry table (which is quadratic on a merge-every-pair
		// sweep like `replace FILE "\n" "" --all`).
		let entryCursor = 0;
		for (const span of matcher(flat)) {
			const start = mapFlatOffset(span.start, entries, entryCursor);
			const end = mapFlatOffset(span.end, entries, start.entryIndex);
			entryCursor = end.entryIndex;
			out.push({
				startBlockId: start.blockId,
				startOffset: start.local,
				endBlockId: end.blockId,
				endOffset: end.local,
				text: span.text,
			});
		}
	}
	return { matches: out, ...normalization };
}

/** Maximal runs of adjacent paragraph blocks in document order. A table or
 *  section-break block ends the current run (a match can't span structure),
 *  but each table cell contributes its own runs — a match can span
 *  consecutive paragraphs inside one cell, never across a cell wall. */
function consecutiveParagraphGroups(blocks: Block[]): Paragraph[][] {
	const groups: Paragraph[][] = [];
	let current: Paragraph[] = [];
	const flush = (): void => {
		if (current.length > 0) {
			groups.push(current);
			current = [];
		}
	};
	for (const block of blocks) {
		if (block.type === "paragraph") {
			current.push(block);
			continue;
		}
		flush();
		if (block.type === "table") {
			for (const row of block.rows) {
				for (const cell of row.cells) {
					groups.push(...consecutiveParagraphGroups(cell.blocks));
				}
			}
		}
	}
	flush();
	return groups;
}

/** Map a flat-text offset back to (blockId, paragraph-local offset). A boundary
 *  offset (exactly at a paragraph's text end, where the "\n" separator sits) maps
 *  to the END of that paragraph, so a match that stops there doesn't spill into
 *  the next one. `fromIndex` is where the scan resumes — callers pass the
 *  previous hit's `entryIndex`, which is safe because both matcher spans and
 *  the entry table are ascending. */
function mapFlatOffset(
	offset: number,
	entries: { blockId: string; flatStart: number; length: number }[],
	fromIndex: number,
): { blockId: string; local: number; entryIndex: number } {
	for (let index = fromIndex; index < entries.length; index++) {
		const entry = entries[index];
		if (entry && offset <= entry.flatStart + entry.length) {
			return {
				blockId: entry.blockId,
				local: Math.max(0, offset - entry.flatStart),
				entryIndex: index,
			};
		}
	}
	throw new Error(`mapFlatOffset: offset ${offset} out of range`);
}

export type AcrossReplaceSpec = {
	pattern: string;
	replacement: string;
	regex: boolean;
	ignoreCase: boolean;
	exact: boolean;
	all: boolean;
	limit?: number;
	view: FindView;
};

export type AcrossReplaceResult = {
	totalMatches: number;
	/** The matches actually replaced (or that WOULD be, under dryRun), in
	 *  document order. */
	replaced: ParagraphSpanMatch[];
	normalizedQuery?: string;
	normalizationApplied?: NormalizationKind[];
};

/** Find + (optionally) apply one cross-paragraph substitution spec — the
 *  engine behind `replace` with a "\n" (single-shot and --batch, whose loop
 *  rereads between entries so each entry's ids are fresh). Applies in REVERSE
 *  document order: matches are document-ordered and non-overlapping, so
 *  mutating from the back keeps every earlier match's node refs and offsets
 *  valid — even when two matches share a boundary paragraph, the earlier one's
 *  span ends before the later one's rewrite begins. */
export function applyAcrossReplace(
	doc: Body,
	spec: AcrossReplaceSpec,
	{ dryRun = false }: { dryRun?: boolean } = {},
): AcrossReplaceResult {
	const { matches, ...normalization } = findAcrossParagraphs(
		doc,
		spec.pattern,
		{
			regex: spec.regex,
			ignoreCase: spec.ignoreCase,
			view: spec.view,
			exact: spec.exact,
		},
	);
	const replaced = selectMatches(matches, spec);
	if (!dryRun) {
		const expand = replacementExpander(spec);
		for (const match of [...replaced].reverse()) {
			replaceAcrossParagraphs(doc, match, expand(match.text), spec.view);
		}
	}
	return { totalMatches: matches.length, replaced, ...normalization };
}

/** First/all/limit selection, shared by every replace surface: `limit` caps
 *  the sweep, `all` takes every match, the default takes just the first. */
export function selectMatches<MatchShape>(
	matches: MatchShape[],
	options: { all: boolean; limit?: number | undefined },
): MatchShape[] {
	if (options.limit !== undefined) return matches.slice(0, options.limit);
	return options.all ? matches : matches.slice(0, 1);
}

/** The concrete replacement for one match: literal replacements pass through
 *  as-is; regex replacements re-run the pattern against the matched text so
 *  `$1`-style backreferences expand. The pattern compiles ONCE — call the
 *  returned function per match. */
export function replacementExpander(spec: {
	pattern: string;
	replacement: string;
	regex: boolean;
	ignoreCase: boolean;
}): (matchText: string) => string {
	if (!spec.regex) return () => spec.replacement;
	const regex = new RegExp(spec.pattern, spec.ignoreCase ? "i" : "");
	return (matchText) => matchText.replace(regex, spec.replacement);
}

/** Filter for `find` by run-level formatting (the inverse workflow of
 *  `edit --clear`): locate the spans carrying a given highlight/color/style so
 *  an agent can strip or re-style them. `highlight: "any"` matches any
 *  highlight color; a specific name matches that color. */
export type RunFormatFilter = {
	highlight?: string;
	color?: string;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
};

/** Return the span locators of runs whose formatting matches `filter`.
 *  Adjacent matching runs coalesce into one span. Offsets are in the same
 *  accepted-view coordinate space as `findTextSpans`, so results paste straight
 *  into `edit --at <span> --clear …` / `comments add --at`. */
export function findFormattedSpans(
	doc: Body,
	filter: RunFormatFilter,
	view: FindView = "accepted",
): TextMatch[] {
	const out: TextMatch[] = [];
	for (const block of iterateBlocks(doc.blocks)) {
		if (block.type !== "paragraph") continue;
		let offset = 0;
		let spanStart: number | null = null;
		let spanText = "";
		const flush = (end: number): void => {
			if (spanStart !== null && spanText.length > 0) {
				out.push({ blockId: block.id, start: spanStart, end, text: spanText });
			}
			spanStart = null;
			spanText = "";
		};
		for (const run of block.runs) {
			if (run.type !== "text") {
				// A tab/line break can't carry run formatting, so it ends the
				// current span — but it still occupies its offset slot (advance so
				// span offsets stay aligned with `paragraphTextForView`).
				flush(offset);
				offset += runViewText(run, view).length;
				continue;
			}
			if (!isRunVisibleInView(run.trackedChange?.kind, view)) continue;
			if (runMatchesFilter(run, filter)) {
				if (spanStart === null) spanStart = offset;
				spanText += run.text;
			} else {
				flush(offset);
			}
			offset += run.text.length;
		}
		flush(offset);
	}
	return out;
}

function runMatchesFilter(
	run: Extract<Paragraph["runs"][number], { type: "text" }>,
	filter: RunFormatFilter,
): boolean {
	if (filter.bold && run.bold !== true) return false;
	if (filter.italic && run.italic !== true) return false;
	if (filter.underline && !run.underline) return false;
	if (filter.highlight !== undefined) {
		if (!run.highlight) return false;
		if (
			filter.highlight !== "any" &&
			run.highlight.toLowerCase() !== filter.highlight.toLowerCase()
		) {
			return false;
		}
	}
	if (filter.color !== undefined) {
		const runColor = (run.color ?? "").toLowerCase().replace(/^#/, "");
		if (runColor !== filter.color.toLowerCase().replace(/^#/, "")) return false;
	}
	return true;
}

type SpanMatch = { start: number; end: number; text: string };
type Matcher = (paragraphText: string) => SpanMatch[];

function literalMatcher(
	query: string,
	ignoreCase: boolean,
	normalize: boolean,
): Matcher {
	if (query.length === 0) {
		throw new Error("query cannot be empty");
	}
	const needle = ignoreCase ? query.toLowerCase() : query;
	return (paragraphText) => {
		const canonical = normalize
			? normalizeHaystack(paragraphText)
			: paragraphText;
		const haystack = ignoreCase ? canonical.toLowerCase() : canonical;
		const matches: SpanMatch[] = [];
		let cursor = haystack.indexOf(needle);
		while (cursor !== -1) {
			matches.push({
				start: cursor,
				end: cursor + needle.length,
				text: paragraphText.slice(cursor, cursor + needle.length),
			});
			cursor = haystack.indexOf(needle, cursor + needle.length);
		}
		return matches;
	};
}

function regexMatcher(pattern: string, ignoreCase: boolean): Matcher {
	const flags = `g${ignoreCase ? "i" : ""}`;
	const regex = new RegExp(pattern, flags);
	return (paragraphText) => {
		const matches: SpanMatch[] = [];
		regex.lastIndex = 0;
		let result = regex.exec(paragraphText);
		while (result !== null) {
			const matched = result[0];
			if (matched.length === 0) {
				// Avoid an infinite loop on zero-width matches.
				regex.lastIndex += 1;
				result = regex.exec(paragraphText);
				continue;
			}
			matches.push({
				start: result.index,
				end: result.index + matched.length,
				text: matched,
			});
			result = regex.exec(paragraphText);
		}
		return matches;
	};
}

function collectMatches(
	blocks: Block[],
	matcher: Matcher,
	view: FindView,
	out: TextMatch[],
): void {
	for (const block of iterateBlocks(blocks)) {
		if (block.type !== "paragraph") continue;
		const paragraphText = paragraphTextForView(block, view);
		for (const span of matcher(paragraphText)) {
			const match: TextMatch = {
				blockId: block.id,
				start: span.start,
				end: span.end,
				text: span.text,
			};
			const overlaps = trackedChangesOverlapping(
				block,
				span.start,
				span.end,
				view,
			);
			if (overlaps.length > 0) match.trackedChanges = overlaps;
			out.push(match);
		}
	}
}

function paragraphTextForView(paragraph: Paragraph, view: FindView): string {
	let out = "";
	for (const run of paragraph.runs) out += runViewText(run, view);
	return out;
}

/** The text a run contributes in the given view — the AST-side counterpart of
 *  `inlineMarkerWidth` in run-ops.ts. Visible text is itself; a visible tab or
 *  line break renders as "\t"/"\n" in `read`, so `find` matches them the same
 *  way; everything else (images, math, page/column breaks, hidden runs) is "".
 *  Every offset walker here derives both its haystack text and its offset
 *  widths from this one function, so they can't disagree. */
function runViewText(run: Paragraph["runs"][number], view: FindView): string {
	if (run.type === "text") {
		return isRunVisibleInView(run.trackedChange?.kind, view) ? run.text : "";
	}
	if (run.type === "tab") {
		return isRunVisibleInView(run.trackedChange?.kind, view) ? "\t" : "";
	}
	if (run.type === "break") {
		return run.kind === "line" &&
			isRunVisibleInView(run.trackedChange?.kind, view)
			? "\n"
			: "";
	}
	return "";
}

function isRunVisibleInView(
	kind: TrackedChange["kind"] | undefined,
	view: FindView,
): boolean {
	if (view === "current") return true;
	if (view === "accepted") return kind !== "del" && kind !== "moveFrom";
	return kind !== "ins" && kind !== "moveTo";
}

function normalizeQuery(query: string): {
	normalized: string;
	applied: NormalizationKind[];
} {
	const applied: NormalizationKind[] = [];
	let result = query;
	const stripped = stripBalancedMarkdownEmphasis(result);
	if (stripped !== result) {
		applied.push("strip-md-emphasis");
		result = stripped;
	}
	const quoteNormalized = normalizeQuotes(result);
	if (quoteNormalized !== result) {
		applied.push("smart-quotes");
		result = quoteNormalized;
	}
	const dashNormalized = normalizeDashes(result);
	if (dashNormalized !== result) {
		applied.push("dashes");
		result = dashNormalized;
	}
	const whitespaceNormalized = normalizeWhitespace(result);
	if (whitespaceNormalized !== result) {
		applied.push("whitespace");
		result = whitespaceNormalized;
	}
	const bulletNormalized = normalizeBullets(result);
	if (bulletNormalized !== result) {
		applied.push("bullets");
		result = bulletNormalized;
	}
	return { normalized: result, applied };
}

function normalizeHaystack(text: string): string {
	return normalizeBullets(
		normalizeWhitespace(normalizeDashes(normalizeQuotes(text))),
	);
}

function normalizeQuotes(text: string): string {
	return text.replace(/[‘’]/g, "'").replace(/[“”]/g, '"');
}

function normalizeDashes(text: string): string {
	// Em-dash (U+2014) and en-dash (U+2013) → hyphen. Both are 1 character ⇒
	// the canonical form is the same length, so match offsets line up with
	// original-text offsets without an index map. The double-hyphen `--` is
	// intentionally NOT normalized (ambiguous: subtraction, CLI flags).
	return text.replace(/[–—]/g, "-");
}

function normalizeWhitespace(text: string): string {
	// Horizontal whitespace variants → a plain space, so a typed space matches a
	// TAB or a non-breaking / typographic space in the document (résumé
	// placeholder lines separate their fields with tabs). Each is one character
	// wide, so match offsets line up with the original text without an index map.
	// Newlines are deliberately left alone — they carry line/paragraph structure,
	// not intra-line spacing.
	return text.replace(/[\t\u00a0\u1680\u2000-\u200a\u202f\u205f\u3000]/g, " ");
}

function normalizeBullets(text: string): string {
	// Interchangeable bullet / separator dots → a canonical bullet (U+2022), so a
	// pattern copied with one glyph matches a document authored with another. All
	// one character wide, so offsets are preserved.
	return text.replace(/[·•‣∙▪◦]/g, "•");
}

/** Strip balanced markdown emphasis markers around non-whitespace content:
 *  `**X**`, `__X__`, `*X*`, `_X_`, `` `X` ``. Conservative: a marker is only
 *  stripped when it has a matching closer with non-whitespace at both inner
 *  boundaries (markdown emphasis grammar) — preserves "5 * 3", `snake_case`,
 *  unmatched asterisks, etc. */
function stripBalancedMarkdownEmphasis(text: string): string {
	const patterns: RegExp[] = [
		/\*\*(\S(?:.*?\S)?)\*\*/g,
		/__(\S(?:.*?\S)?)__/g,
		/`(\S(?:.*?\S)?)`/g,
		/(?<![A-Za-z0-9_])\*(\S(?:.*?\S)?)\*(?![A-Za-z0-9_])/g,
		/(?<![A-Za-z0-9_])_(\S(?:.*?\S)?)_(?![A-Za-z0-9_])/g,
	];
	let result = text;
	for (const pattern of patterns) {
		result = result.replace(pattern, "$1");
	}
	return result;
}

function trackedChangesOverlapping(
	paragraph: Paragraph,
	start: number,
	end: number,
	view: FindView,
): TrackedChange[] {
	const seen = new Set<string>();
	const out: TrackedChange[] = [];
	let offset = 0;
	for (const run of paragraph.runs) {
		// Advance the offset by each run's view width (tabs/line breaks count as
		// one, exactly as in `paragraphTextForView`) so overlap tests line up.
		const length = runViewText(run, view).length;
		if (length === 0) continue;
		const runStart = offset;
		const runEnd = offset + length;
		offset = runEnd;
		if (runEnd <= start || runStart >= end) continue;
		const change =
			run.type === "text" || run.type === "tab" || run.type === "break"
				? run.trackedChange
				: undefined;
		if (!change) continue;
		if (seen.has(change.id)) continue;
		seen.add(change.id);
		out.push(change);
	}
	return out;
}
