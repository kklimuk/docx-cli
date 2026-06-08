import {
	type FindView,
	findFormattedSpans,
	findTextSpans,
	type RunFormatFilter,
	type TextMatch,
} from "@core/find";
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
  docx find FILE --highlight COLOR|any [--all]   # find by formatting (no QUERY)

Positional:
  QUERY             literal substring (or regex if --regex). Omit it when using
                    a formatting filter below.

Formatting filters (alternative to QUERY — locate runs by formatting, the
inverse of \`edit --clear\`; pair with \`edit --at <span> --clear\`):
  --highlight C     runs highlighted color C (a name like "yellow", or "any")
  --color HEX       runs with text color HEX (e.g. FF0000)
  --bold            bold runs
  --italic          italic runs
  --underline       underlined runs

Options:
  --regex           treat QUERY as a JavaScript regular expression
  --ignore-case     case-insensitive match
  --all             return every match (default: just the first)
  --nth N           return only the Nth match (0-indexed)
  --current         search the raw concatenation (both ins and del text)
  --baseline        search the pre-change text (skip ins/moveTo)
                    (--current and --baseline are mutually exclusive; default:
                    accepted document (skip del/moveFrom) — matches
                    "docx read" / "docx wc" / "docx comments add")
  --exact           disable query normalization (no markdown-emphasis stripping,
                    no smart/straight quote or em/en-dash equivalence)
  --json            emit the full match objects as JSON (default: bare locators)
  -h, --help        show this help

Within-paragraph matches only — cross-paragraph ranges aren't supported
yet. Searches the concatenated text of each paragraph in document order,
including paragraphs nested in table cells (locators look like
tT:rRcC:pK:S-E for those).

By default the query is normalized: balanced markdown emphasis around
non-whitespace (**X**, __X__, *X*, \`X\`) is stripped; smart quotes match
straight quotes; em-dash and en-dash match the hyphen. Pass --exact to
match the raw query verbatim. --regex is always verbatim.

Output:
  Default: the matched span locators (e.g. p3:5-8), one per line — feed them
  straight into another command's --at. No matches prints nothing (exit 0).
  --json: { totalMatches, query, view, matches:[{locator, blockId, start, end,
  text, …}], normalizedQuery? } (no envelope). Errors print {code, error,
  hint?} with a nonzero exit. Notation: offsets are 0-based, end-exclusive.

Examples:
  docx find doc.docx "fox"
  docx find doc.docx "Action Item:" --all
  docx find doc.docx "TODO|FIXME" --regex --ignore-case
  docx find doc.docx --highlight yellow --all     # every yellow-highlighted span
  docx find doc.docx --highlight any --all | while read s; do \\
    docx edit doc.docx --at "$s" --clear highlight; done
  docx comments add doc.docx --at "$(docx find doc.docx fox | head -1)" --text "..."
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
	const query = parsed.positionals[1];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	// Formatting filters (highlight/color/bold/italic/underline) are an
	// alternative to a text QUERY — they locate runs by their formatting (the
	// inverse of `edit --clear`).
	const formatFilter: RunFormatFilter = {};
	if (parsed.values.highlight !== undefined)
		formatFilter.highlight = parsed.values.highlight as string;
	if (parsed.values.color !== undefined)
		formatFilter.color = parsed.values.color as string;
	if (parsed.values.bold) formatFilter.bold = true;
	if (parsed.values.italic) formatFilter.italic = true;
	if (parsed.values.underline) formatFilter.underline = true;
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

	let result: ReturnType<typeof findTextSpans> = { matches: [] };
	if (hasFormatFilter) {
		result = {
			matches: findFormattedSpans(document.body, formatFilter, findView),
		};
	} else if (query != null) {
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

	const matches = selected.map((match) => ({
		locator: `${match.blockId}:${match.start}-${match.end}`,
		...match,
	}));

	if (parsed.values.json) {
		await respond({
			totalMatches: allMatches.length,
			query,
			regex: useRegex,
			ignoreCase,
			view: findView,
			matches,
			...(result.normalizedQuery !== undefined
				? {
						normalizedQuery: result.normalizedQuery,
						normalizationApplied: result.normalizationApplied,
					}
				: {}),
		});
		return EXIT.OK;
	}

	// Text-first default: the locators, one per line, ready to paste into --at.
	if (matches.length > 0) {
		await writeStdout(`${matches.map((match) => match.locator).join("\n")}\n`);
	}
	return EXIT.OK;
}
