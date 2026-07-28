import type { Document } from "@core";
import {
	type FindView,
	findAcrossParagraphs,
	findFormattedSpans,
	findTextSpans,
	type ParagraphSpanMatch,
	type RunFormatFilter,
	type TextMatch,
} from "@core/find";
import { spanLocator } from "../parse-helpers";
import type { ErrorCode } from "../respond";

/** Execute one validated find request against an already-open document. Text,
 * formatting, and cross-paragraph queries share one result contract so the
 * single-shot and JSONL batch surfaces cannot drift. */
export function executeFindQuery(
	document: Document,
	request: FindQueryRequest,
): FindQueryResult {
	let matches: FindMatchPayload[];
	let normalizationFields: NormalizationPayload = {};
	try {
		if (request.query?.includes("\n")) {
			const result = findAcrossParagraphs(document.body, request.query, {
				regex: request.regex,
				ignoreCase: request.ignoreCase,
				view: request.view,
				exact: request.exact,
			});
			matches = result.matches.map((match) => ({
				locator: spanLocator(match),
				...match,
			}));
			normalizationFields = normalizationPayload(result);
		} else if (request.formatFilter) {
			matches = findFormattedSpans(
				document.body,
				request.formatFilter,
				request.view,
			).map((match) => ({ locator: spanLocator(match), ...match }));
		} else if (request.query !== undefined) {
			const result = findTextSpans(document.body, request.query, {
				regex: request.regex,
				ignoreCase: request.ignoreCase,
				view: request.view,
				exact: request.exact,
			});
			matches = result.matches.map((match) => ({
				locator: spanLocator(match),
				...match,
			}));
			normalizationFields = normalizationPayload(result);
		} else {
			throw new FindQueryError("USAGE", "Missing query or formatting filter");
		}
	} catch (error) {
		if (error instanceof FindQueryError) throw error;
		const message = error instanceof Error ? error.message : String(error);
		throw new FindQueryError("USAGE", `Invalid query: ${message}`);
	}

	let selected = matches;
	if (request.nth !== undefined) {
		const single = matches[request.nth];
		if (!single) {
			throw new FindQueryError(
				"MATCH_NOT_FOUND",
				`Only ${matches.length} match(es); nth ${request.nth} is out of range`,
			);
		}
		selected = [single];
	}

	return {
		totalMatches: matches.length,
		...(request.query !== undefined ? { query: request.query } : {}),
		...(request.formatFilter ? { filter: request.formatFilter } : {}),
		regex: request.regex,
		ignoreCase: request.ignoreCase,
		view: request.view,
		matches: selected,
		...normalizationFields,
	};
}

function normalizationPayload(result: {
	normalizedQuery?: string;
	normalizationApplied?: string[];
}): NormalizationPayload {
	return result.normalizedQuery !== undefined
		? {
				normalizedQuery: result.normalizedQuery,
				normalizationApplied: result.normalizationApplied,
			}
		: {};
}

export class FindQueryError extends Error {
	constructor(
		public code: ErrorCode,
		message: string,
	) {
		super(message);
		this.name = "FindQueryError";
	}
}

export type FindQueryRequest = {
	query?: string;
	formatFilter?: RunFormatFilter;
	regex: boolean;
	ignoreCase: boolean;
	exact: boolean;
	view: FindView;
	nth?: number;
};

export type FindQueryResult = {
	totalMatches: number;
	query?: string;
	filter?: RunFormatFilter;
	regex: boolean;
	ignoreCase: boolean;
	view: FindView;
	matches: FindMatchPayload[];
	normalizedQuery?: string;
	normalizationApplied?: string[];
};

type FindMatchPayload =
	| ({ locator: string } & TextMatch)
	| ({ locator: string } & ParagraphSpanMatch);

type NormalizationPayload = Pick<
	FindQueryResult,
	"normalizedQuery" | "normalizationApplied"
>;
