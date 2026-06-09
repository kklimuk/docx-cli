import {
	findFormattedSpans,
	findTextSpans,
	type RunFormatFilter,
	type TextMatch,
} from "@core/find";
import { resolveView } from "../parse-helpers";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	tryParseArgs,
	writeStderr,
	writeStdout,
} from "../respond";

const HELP = `docx find — locate text spans and return their locators

Usage:
  docx find FILE QUERY [options]
  docx find FILE --highlight [COLOR]             # find by formatting (no QUERY)

Positional:
  QUERY             literal substring (or regex if --regex). Omit it when using
                    a formatting filter below.

Formatting filters (alternative to QUERY — locate runs by formatting, the
inverse of \`edit --clear\`; pair with \`edit --at <span> --clear\`):
  --highlight [C]   runs highlighted color C (a name like "yellow", or "any").
                    Bare --highlight (no value) means any color. Returns each
                    highlighted placeholder's FULL span — paste it straight into
                    \`edit --at\` instead of hand-counting offsets from a text match.
  --color HEX       runs with text color HEX (e.g. FF0000)
  --bold            bold runs
  --italic          italic runs
  --underline       underlined runs

Options:
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

Within-paragraph matches only — cross-paragraph ranges aren't supported
yet. Searches the concatenated text of each paragraph in document order,
including paragraphs nested in table cells (locators look like
tT:rRcC:pK:S-E for those).

By default the query is normalized: balanced markdown emphasis around
non-whitespace (**X**, __X__, *X*, \`X\`) is stripped; smart quotes match
straight quotes; em-dash and en-dash match the hyphen. Pass --exact to
match the raw query verbatim. --regex is always verbatim.

Output:
  Default: EVERY matched span locator (e.g. p3:5-8), one per line — feed them
  straight into another command's --at (or a --batch). No matches prints
  nothing to stdout and "no matches" to stderr (exit 0) — so an empty result is
  unambiguous. Use --nth N for a single match, or pipe to "head -1".
  --json: { totalMatches, query, view, matches:[{locator, blockId, start, end,
  text, …}], normalizedQuery? } (no envelope). Errors print {code, error,
  hint?} with a nonzero exit. Notation: offsets are 0-based, end-exclusive.

Examples:
  docx find doc.docx "fox"                         # every match, one per line
  docx find doc.docx "TODO|FIXME" --regex --ignore-case
  docx find doc.docx --highlight yellow            # every yellow-highlighted span
  docx find doc.docx --highlight any | while read s; do \\
    docx edit doc.docx --at "$s" --clear highlight; done
  docx comments add doc.docx --at "$(docx find doc.docx fox | head -1)" --text "..."
`;

/** `--highlight` with no value means "any color" — the common cleanup intent
 *  ("find every highlight, regardless of color"). `--highlight` is a string flag,
 *  so a bare `--highlight` (at the end, or before another flag) would otherwise be
 *  a parse error; rewrite it to `--highlight any` before parsing. `--highlight=`
 *  (explicit empty) is handled too. */
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
	const exact = Boolean(parsed.values.exact);
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
	} else {
		// Default: every match. Returning only the first silently hid the rest —
		// a weak agent that didn't know to pass --all fell back to one find per
		// item. `--all` is now the default; `--nth N` selects a single match and
		// `| head -1` still grabs the first. (`--all` is kept as an accepted no-op.)
		selected = allMatches;
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
	// On zero matches stdout stays empty (so `find … | while read` is clean), but
	// we print an explicit "no matches" to STDERR — otherwise empty output reads
	// as "did it even run?" and (weak) agents re-run the query to be sure.
	if (matches.length > 0) {
		await writeStdout(`${matches.map((match) => match.locator).join("\n")}\n`);
	} else {
		await writeStderr("no matches\n");
	}
	return EXIT.OK;
}
