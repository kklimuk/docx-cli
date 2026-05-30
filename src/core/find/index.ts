import { type Body, iterateBlocks } from "../ast/document/body";
import type { Block, Paragraph, TrackedChange } from "../ast/types";

export {
	replaceSpanInParagraph,
	type Span,
	type TrackedReplaceOptions,
} from "./replace-span";

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

export type NormalizationKind = "strip-md-emphasis" | "smart-quotes" | "dashes";

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
	const exact = options.exact ?? false;
	const useRegex = options.regex ?? false;

	let effectiveQuery = query;
	const applied: NormalizationKind[] = [];
	if (!useRegex && !exact) {
		const norm = normalizeQuery(query);
		effectiveQuery = norm.normalized;
		applied.push(...norm.applied);
	}

	const matcher = useRegex
		? regexMatcher(query, options.ignoreCase ?? false)
		: literalMatcher(effectiveQuery, options.ignoreCase ?? false, !exact);
	const out: TextMatch[] = [];
	collectMatches(doc.blocks, matcher, view, out);

	const result: FindResult = { matches: out };
	if (applied.length > 0) {
		result.normalizedQuery = effectiveQuery;
		result.normalizationApplied = applied;
	}
	return result;
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
	for (const run of paragraph.runs) {
		if (run.type !== "text") continue;
		if (!isRunVisibleInView(run.trackedChange?.kind, view)) continue;
		out += run.text;
	}
	return out;
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
	return { normalized: result, applied };
}

function normalizeHaystack(text: string): string {
	return normalizeDashes(normalizeQuotes(text));
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
		if (run.type !== "text") continue;
		if (!isRunVisibleInView(run.trackedChange?.kind, view)) continue;
		const length = run.text.length;
		const runStart = offset;
		const runEnd = offset + length;
		offset = runEnd;
		if (runEnd <= start || runStart >= end) continue;
		const change = run.trackedChange;
		if (!change) continue;
		if (seen.has(change.id)) continue;
		seen.add(change.id);
		out.push(change);
	}
	return out;
}
