import { type FindView, findTextSpans, type TextMatch } from "@core/find";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx find — locate text spans and return their locators

Usage:
  docx find FILE QUERY [options]

Required positional:
  QUERY             literal substring (or regex if --regex)

Options:
  --regex           treat QUERY as a JavaScript regular expression
  --ignore-case     case-insensitive match
  --all             return every match (default: just the first)
  --nth N           return only the Nth match (0-indexed)
  --current         search the raw concatenation (both ins and del text)
  --baseline        search the pre-change text (skip ins/moveTo)
                    default: accepted document (skip del/moveFrom) — matches
                    "docx read" / "docx wc" / "docx comments add"
  --exact           disable query normalization (no markdown-emphasis stripping,
                    no smart/straight quote or em/en-dash equivalence)
  -h, --help        show this help

Within-paragraph matches only — cross-paragraph ranges aren't supported
yet. Searches the concatenated text of each paragraph in document order,
including paragraphs nested in table cells (locators look like
tT:rRcC:pK:S-E for those).

By default the query is normalized: balanced markdown emphasis around
non-whitespace (**X**, __X__, *X*, \`X\`) is stripped; smart quotes match
straight quotes; em-dash and en-dash match the hyphen. Pass --exact to
match the raw query verbatim. --regex is always verbatim.

Examples:
  docx find doc.docx "fox"
  docx find doc.docx "Action Item:" --all
  docx find doc.docx "TODO|FIXME" --regex --ignore-case
  docx comments add doc.docx --range "$(docx find doc.docx fox | jq -r .matches[0].locator)" --text "..."
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			regex: { type: "boolean" },
			"ignore-case": { type: "boolean" },
			all: { type: "boolean" },
			nth: { type: "string" },
			current: { type: "boolean" },
			baseline: { type: "boolean" },
			exact: { type: "boolean" },
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
	const query = parsed.positionals[1];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);
	if (query == null) return fail("USAGE", "Missing QUERY argument", HELP);

	const ignoreCase = Boolean(parsed.values["ignore-case"]);
	const useRegex = Boolean(parsed.values.regex);
	const wantAll = Boolean(parsed.values.all);
	const wantCurrent = Boolean(parsed.values.current);
	const wantBaseline = Boolean(parsed.values.baseline);
	const exact = Boolean(parsed.values.exact);
	if (wantCurrent && wantBaseline) {
		return fail("USAGE", "--current and --baseline are mutually exclusive");
	}
	const findView: FindView = wantCurrent
		? "current"
		: wantBaseline
			? "baseline"
			: "accepted";
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

	let result: ReturnType<typeof findTextSpans>;
	try {
		result = findTextSpans(document.body, query, {
			regex: useRegex,
			ignoreCase,
			view: findView,
			exact,
		});
	} catch (matcherError) {
		const message =
			matcherError instanceof Error
				? matcherError.message
				: String(matcherError);
		return fail("USAGE", `Invalid query: ${message}`);
	}

	const allMatches = result.matches;
	let selected: TextMatch[];
	if (nth !== undefined) {
		const single = allMatches[nth];
		if (!single) {
			return fail(
				"MATCH_NOT_FOUND",
				`Only ${allMatches.length} match(es); --nth ${nth} is out of range`,
			);
		}
		selected = [single];
	} else if (wantAll) {
		selected = allMatches;
	} else {
		selected = allMatches.slice(0, 1);
	}

	await respond({
		ok: true,
		operation: "find",
		path,
		query,
		regex: useRegex,
		ignoreCase,
		view: findView,
		totalMatches: allMatches.length,
		matches: selected.map((match) => ({
			locator: `${match.blockId}:${match.start}-${match.end}`,
			...match,
		})),
		...(result.normalizedQuery !== undefined
			? {
					normalizedQuery: result.normalizedQuery,
					normalizationApplied: result.normalizationApplied,
				}
			: {}),
	});
	return EXIT.OK;
}
