import type { RunFormatFilter } from "@core/find";
import {
	decodeInlineEscapes,
	rejectBatchOnlyFlags,
	resolveView,
} from "../parse-helpers";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	tryParseArgs,
	writeStderr,
	writeStdout,
} from "../respond";
import { runFindBatch } from "./batch";
import {
	executeFindQuery,
	FindQueryError,
	type FindQueryResult,
} from "./query";

const HELP = `docx find — locate content and return its locator

Usage:
  docx find FILE QUERY [options]
  docx find FILE --highlight [COLOR]             # find by formatting (no QUERY)
  docx find FILE --batch FILE.jsonl              # many queries, one read
  docx find FILE --batch -                       # read JSONL from stdin

Examples:
  docx find doc.docx "fox"                         # every match, one per line
  docx find doc.docx "TODO|FIXME" --regex --ignore-case
  docx find doc.docx --highlight yellow            # every yellow-highlighted span
  # queries.jsonl:
  #   {"query":"TODO|FIXME","regex":true,"ignoreCase":true}
  #   {"query":"Clause A\\nClause B","nth":0}
  #   {"highlight":"yellow"}
  docx find doc.docx --batch queries.jsonl --json
  # remove every highlight:
  docx find doc.docx --highlight any | while read span; do
    docx edit doc.docx --at "$span" --clear highlight; done
  docx comments add doc.docx --at "$(docx find doc.docx fox | head -1)" --text "..."

Positional:
  QUERY             literal substring (or regex if --regex). Omit it when using
                    a formatting filter or --batch.

Formatting filters (alternative to QUERY — locate runs by formatting, the
inverse of \`edit --clear\`; pair with \`edit --at <span> --clear\`):
  --highlight [C]   runs highlighted with color C (a name like "yellow", or
                    "any"). Bare --highlight (no value) means any color.
                    Returns each highlighted stretch's FULL span.
  --color HEX       runs with text color HEX (e.g. FF0000)
  --bold            bold runs
  --italic          italic runs
  --underline       underlined runs

General options:
  --regex           treat QUERY as a JavaScript regular expression
  --ignore-case     case-insensitive match
  --nth N           return only the Nth match (0-indexed). By default EVERY
                    match is returned, one locator per line — pipe to a batch
                    or loop. (--all is accepted but redundant: all is default.)
  --current         search the raw concatenation (both ins and del text)
  --baseline        search the pre-change text (skip ins/moveTo)
                    (--current and --baseline are mutually exclusive; default:
                    accepted document (skip del/moveFrom) — matches
                    "docx read" / "docx wc" / "docx comments add")
  --exact           disable query normalization (no markdown-emphasis stripping,
                    no smart/straight quote or em/en-dash equivalence)
  --json            emit the full match objects as JSON (default: bare locators)
  -h, --help        show this help

Batch (--batch PATH | -):
  Evaluate many independent queries against one document read. Each JSONL line
  uses {"query": …} plus optional "regex"/"ignoreCase"/"nth"/"exact"/
  "current"/"baseline", OR formatting keys "highlight"/"color"/"bold"/
  "italic"/"underline". Default output flattens locators in entry order;
  --json keeps request boundaries as {batch:[{totalMatches,matches,…}, …]}.
  A zero-match entry is successful (and named on stderr in text mode); an invalid
  query or out-of-range "nth" fails the batch with its entry number.

Output:
  Default: EVERY matched span locator (e.g. p3:5-8), one per line — feed them
  straight into another command's --at (or a --batch). No matches prints
  nothing to stdout and "no matches" to stderr (exit 0) — so an empty result is
  unambiguous. Use --nth N for a single match, or pipe to "head -1".
  --json: { totalMatches, query, view, matches:[{locator, blockId, start, end,
  text, …}], normalizedQuery? } (no envelope). Errors print {code, error,
  hint?} with a nonzero exit. Notation: offsets are 0-based, end-exclusive.
`;

/** `--highlight` with no value means "any color" — the common cleanup intent
 * ("find every highlight, regardless of color"). */
function withBareHighlightAsAny(args: string[]): string[] {
	const out: string[] = [];
	for (let index = 0; index < args.length; index++) {
		const arg = args[index];
		if (arg === "--highlight") {
			const next = args[index + 1];
			out.push(arg);
			if (next === undefined || next.startsWith("-")) out.push("any");
			continue;
		}
		if (arg === "--highlight=") {
			out.push("--highlight=any");
			continue;
		}
		if (arg !== undefined) out.push(arg);
	}
	return out;
}

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		withBareHighlightAsAny(args),
		{
			batch: { type: "string" },
			regex: { type: "boolean" },
			"ignore-case": { type: "boolean" },
			all: { type: "boolean" },
			nth: { type: "string" },
			current: { type: "boolean" },
			baseline: { type: "boolean" },
			exact: { type: "boolean" },
			highlight: { type: "string" },
			color: { type: "string" },
			bold: { type: "boolean" },
			italic: { type: "boolean" },
			underline: { type: "boolean" },
			json: { type: "boolean" },
			help: { type: "boolean", short: "h" },
		},
		HELP,
	);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const batchSource = parsed.values.batch as string | undefined;
	if (batchSource !== undefined) {
		if (parsed.positionals.length > 1) {
			return fail(
				"USAGE",
				"--batch reads queries from JSONL; don't also pass a QUERY positional",
				HELP,
			);
		}
		const conflict = await rejectBatchOnlyFlags(
			parsed.values,
			BATCH_ENTRY_ONLY_FLAGS,
			"query",
			"Put query/filter/view/nth fields on each JSONL line. Only --json stays command-wide.",
		);
		if (conflict !== undefined) return conflict;
		return runFindBatch(path, batchSource, Boolean(parsed.values.json));
	}

	const query = decodeInlineEscapes(parsed.positionals[1]);
	const formatFilter = readFormatFilter(parsed.values);
	const hasFormatFilter = Object.keys(formatFilter).length > 0;
	if (query == null && !hasFormatFilter) {
		return fail(
			"USAGE",
			"Missing QUERY (or a --highlight/--color/--bold/--italic/--underline filter)",
			HELP,
		);
	}
	if (query != null && hasFormatFilter) {
		return fail(
			"USAGE",
			"Pass a text QUERY or formatting filters (--highlight/--color/...), not both",
			HELP,
		);
	}

	const findView = resolveView(parsed.values);
	if (!findView) {
		return fail("USAGE", "--current and --baseline are mutually exclusive");
	}
	const nthRaw = parsed.values.nth as string | undefined;
	const nth = nthRaw === undefined ? undefined : Number(nthRaw);
	if (nth !== undefined && (!Number.isInteger(nth) || nth < 0)) {
		return fail(
			"USAGE",
			`--nth must be a non-negative integer, got "${nthRaw}"`,
		);
	}

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	let result: FindQueryResult;
	try {
		result = executeFindQuery(document, {
			...(query !== undefined ? { query } : {}),
			...(hasFormatFilter ? { formatFilter } : {}),
			regex: Boolean(parsed.values.regex),
			ignoreCase: Boolean(parsed.values["ignore-case"]),
			exact: Boolean(parsed.values.exact),
			view: findView,
			...(nth !== undefined ? { nth } : {}),
		});
	} catch (error) {
		if (error instanceof FindQueryError) {
			return fail(error.code, error.message);
		}
		throw error;
	}

	if (parsed.values.json) {
		await respond(result);
		return EXIT.OK;
	}
	if (result.matches.length > 0) {
		await writeStdout(
			`${result.matches.map((match) => match.locator).join("\n")}\n`,
		);
	} else {
		await writeStderr("no matches\n");
	}
	return EXIT.OK;
}

function readFormatFilter(values: Record<string, unknown>): RunFormatFilter {
	const filter: RunFormatFilter = {};
	if (values.highlight !== undefined)
		filter.highlight = values.highlight as string;
	if (values.color !== undefined) filter.color = values.color as string;
	if (values.bold) filter.bold = true;
	if (values.italic) filter.italic = true;
	if (values.underline) filter.underline = true;
	return filter;
}

const BATCH_ENTRY_ONLY_FLAGS = [
	"regex",
	"ignore-case",
	"all",
	"nth",
	"current",
	"baseline",
	"exact",
	"highlight",
	"color",
	"bold",
	"italic",
	"underline",
] as const;
