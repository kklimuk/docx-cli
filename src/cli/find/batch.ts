import type { RunFormatFilter } from "@core/find";
import {
	parseBatchInteger,
	readJsonlObjects,
	requireBatchBoolean,
	resolveView,
} from "../parse-helpers";
import {
	EntryError,
	EXIT,
	fail,
	openOrFail,
	respond,
	writeStderr,
	writeStdout,
} from "../respond";
import {
	executeFindQuery,
	FindQueryError,
	type FindQueryRequest,
	type FindQueryResult,
} from "./query";

/** Evaluate many independent queries against one immutable document read. Each
 * JSONL entry owns its selector/view/options; unlike replace batch there is no
 * sequential mutation or reread between entries. */
export async function runFindBatch(
	filePath: string,
	batchSource: string,
	json: boolean,
): Promise<number> {
	let rawEntries: Record<string, unknown>[];
	try {
		rawEntries = await readJsonlObjects(batchSource);
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail("USAGE", `Failed to read batch: ${message}`);
	}
	if (rawEntries.length === 0) return fail("USAGE", "Batch file is empty");

	let requests: FindQueryRequest[];
	try {
		requests = rawEntries.map((raw, index) => validateEntry(raw, index));
	} catch (error) {
		if (error instanceof EntryError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	const results: FindQueryResult[] = [];
	for (let index = 0; index < requests.length; index++) {
		const request = requests[index];
		if (!request) continue;
		try {
			results.push(executeFindQuery(document, request));
		} catch (error) {
			if (error instanceof FindQueryError) {
				return fail(error.code, `entry ${index}: ${error.message}`);
			}
			throw error;
		}
	}

	if (json) {
		await respond({ batch: results });
		return EXIT.OK;
	}

	const locators = results.flatMap((result) =>
		result.matches.map((match) => match.locator),
	);
	if (locators.length > 0) await writeStdout(`${locators.join("\n")}\n`);
	const misses = results
		.map((result, index) => ({ index, count: result.matches.length }))
		.filter((result) => result.count === 0);
	if (misses.length > 0) {
		await writeStderr(
			`${misses.map((result) => `entry ${result.index}: no matches`).join("\n")}\n`,
		);
	}
	return EXIT.OK;
}

function validateEntry(
	raw: Record<string, unknown>,
	index: number,
): FindQueryRequest {
	const query = raw.query;
	if (
		query !== undefined &&
		(typeof query !== "string" || query.length === 0)
	) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: "query" must be a non-empty string`,
		);
	}

	const formatFilter = readFormatFilter(raw, index);
	const hasFormatFilter = Object.keys(formatFilter).length > 0;
	if (query === undefined && !hasFormatFilter) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: provide "query" or a formatting filter (highlight/color/bold/italic/underline)`,
		);
	}
	if (query !== undefined && hasFormatFilter) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: pass "query" or formatting filters, not both`,
		);
	}

	const regex = requireBatchBoolean(raw, index, "regex");
	const ignoreCase =
		requireBatchBoolean(raw, index, "ignoreCase") ||
		requireBatchBoolean(raw, index, "ignore-case");
	const exact = requireBatchBoolean(raw, index, "exact");
	// `all` is find's default and has no per-entry effect, but an agent that
	// carried it over from `replace --batch` still gets it type-checked.
	requireBatchBoolean(raw, index, "all");
	if (hasFormatFilter && (regex || ignoreCase || exact)) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: regex/ignoreCase/exact apply to a text query, not formatting filters`,
		);
	}

	const view = resolveView({
		current: requireBatchBoolean(raw, index, "current"),
		baseline: requireBatchBoolean(raw, index, "baseline"),
	});
	if (view === null) {
		throw new EntryError(
			"USAGE",
			`entry ${index}: "current" and "baseline" are mutually exclusive`,
		);
	}

	const nthValue = parseBatchInteger(raw, index, "nth", 0);
	if (nthValue !== undefined && typeof nthValue !== "number") {
		throw new EntryError("USAGE", nthValue.error, nthValue.hint);
	}
	const nth = nthValue;

	return {
		...(typeof query === "string" ? { query } : {}),
		...(hasFormatFilter ? { formatFilter } : {}),
		regex,
		ignoreCase,
		exact,
		view,
		...(nth !== undefined ? { nth } : {}),
	};
}

function readFormatFilter(
	raw: Record<string, unknown>,
	index: number,
): RunFormatFilter {
	const filter: RunFormatFilter = {};
	if (raw.highlight !== undefined) {
		if (raw.highlight === true) filter.highlight = "any";
		else if (typeof raw.highlight === "string") {
			filter.highlight = raw.highlight;
		} else {
			throw new EntryError(
				"USAGE",
				`entry ${index}: "highlight" must be a color string or true (any color)`,
			);
		}
	}
	if (raw.color !== undefined) {
		if (typeof raw.color !== "string") {
			throw new EntryError("USAGE", `entry ${index}: "color" must be a string`);
		}
		filter.color = raw.color;
	}
	if (requireBatchBoolean(raw, index, "bold")) filter.bold = true;
	if (requireBatchBoolean(raw, index, "italic")) filter.italic = true;
	if (requireBatchBoolean(raw, index, "underline")) filter.underline = true;
	return filter;
}
