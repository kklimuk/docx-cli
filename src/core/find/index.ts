import type { Block, Doc } from "../ast/types";

export type TextMatch = {
	blockId: string;
	start: number;
	end: number;
	text: string;
};

export type FindOptions = {
	regex?: boolean;
	ignoreCase?: boolean;
};

export function findTextSpans(
	doc: Doc,
	query: string,
	options: FindOptions = {},
): TextMatch[] {
	const matcher = options.regex
		? regexMatcher(query, options.ignoreCase ?? false)
		: literalMatcher(query, options.ignoreCase ?? false);
	const out: TextMatch[] = [];
	collectMatches(doc.blocks, matcher, out);
	return out;
}

type SpanMatch = { start: number; end: number; text: string };
type Matcher = (paragraphText: string) => SpanMatch[];

function literalMatcher(query: string, ignoreCase: boolean): Matcher {
	if (query.length === 0) {
		throw new Error("query cannot be empty");
	}
	const needle = ignoreCase ? query.toLowerCase() : query;
	return (paragraphText) => {
		const haystack = ignoreCase ? paragraphText.toLowerCase() : paragraphText;
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
	out: TextMatch[],
): void {
	for (const block of blocks) {
		if (block.type === "paragraph") {
			const paragraphText = block.runs
				.map((run) => (run.type === "text" ? run.text : ""))
				.join("");
			for (const span of matcher(paragraphText)) {
				out.push({
					blockId: block.id,
					start: span.start,
					end: span.end,
					text: span.text,
				});
			}
			continue;
		}
		if (block.type === "table") {
			for (const row of block.rows) {
				for (const cell of row.cells) {
					collectMatches(cell.blocks, matcher, out);
				}
			}
		}
	}
}
