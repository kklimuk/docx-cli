import { findTextSpans, type TextMatch } from "@core/find";
import { parseArgs } from "util";
import { EXIT, fail, openOrFail, respond, writeStdout } from "../respond";

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
  -h, --help        show this help

Within-paragraph matches only — cross-paragraph ranges aren't supported
yet. Searches the concatenated text of each paragraph in document order,
including paragraphs nested in table cells (locators look like
tT:rRcC:pK:S-E for those).

Examples:
  docx find doc.docx "fox"
  docx find doc.docx "Action Item:" --all
  docx find doc.docx "TODO|FIXME" --regex --ignore-case
  docx comments add doc.docx --range "$(docx find doc.docx fox | jq -r .matches[0].locator)" --text "..."
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				regex: { type: "boolean" },
				"ignore-case": { type: "boolean" },
				all: { type: "boolean" },
				nth: { type: "string" },
				help: { type: "boolean", short: "h" },
			},
		});
	} catch (parseError) {
		const message =
			parseError instanceof Error ? parseError.message : String(parseError);
		return fail("USAGE", message, HELP);
	}

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
	const nthRaw = parsed.values.nth as string | undefined;
	const nth = nthRaw === undefined ? undefined : Number(nthRaw);
	if (nth !== undefined && (!Number.isInteger(nth) || nth < 0)) {
		return fail(
			"USAGE",
			`--nth must be a non-negative integer, got "${nthRaw}"`,
		);
	}

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	let allMatches: TextMatch[];
	try {
		allMatches = findTextSpans(view.doc, query, {
			regex: useRegex,
			ignoreCase,
		});
	} catch (matcherError) {
		const message =
			matcherError instanceof Error
				? matcherError.message
				: String(matcherError);
		return fail("USAGE", `Invalid query: ${message}`);
	}

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
		totalMatches: allMatches.length,
		matches: selected.map((match) => ({
			locator: `${match.blockId}:${match.start}-${match.end}`,
			...match,
		})),
	});
	return EXIT.OK;
}
