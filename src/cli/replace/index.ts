import { resolveAuthor, resolveDate, TrackChanges } from "@core";
import {
	findTextSpans,
	replaceSpanInParagraph,
	type TextMatch,
	type TrackedReplaceOptions,
} from "@core/find";
import { resolveView } from "../parse-helpers";
import {
	EXIT,
	fail,
	openOrFail,
	resolveTracked,
	respond,
	respondAck,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";
import { runReplaceBatch } from "./batch";

const HELP = `docx replace — substitute text spans (sed for docx)

Usage:
  docx replace FILE PATTERN REPLACEMENT [options]
  docx replace FILE --batch FILE.jsonl [options]   # a sed-script, one read
  docx replace FILE --batch -          [options]   # read JSONL from stdin

Options:
  --regex           treat PATTERN as a JavaScript regular expression
  --ignore-case     case-insensitive match
  --all             replace every match (default: just the first)
  --limit N         replace at most N matches (in document order)
  --author NAME     author for tracked changes (default: $DOCX_AUTHOR)
  --track           record substitutions as tracked changes even when the
                    document's track-changes toggle is off (OFF by default)
  --current         operate on the raw concatenation (both ins and del text)
  --baseline        operate on the pre-change text (skip ins/moveTo)
                    (--current and --baseline are mutually exclusive; default:
                    accepted document (skip del/moveFrom) — matches
                    "docx find" / "docx read")
  --exact           disable pattern normalization (no markdown-emphasis stripping,
                    no smart/straight quote or em/en-dash equivalence)
  -o, --output PATH write to PATH instead of overwriting FILE
  --dry-run         report what would change without writing the file
  -v, --verbose     print the success ack JSON (default: a one-line confirmation)
  -h, --help        show this help

Within-paragraph matches only. Run formatting (rPr) on the surrounding text
is preserved; the replacement run inherits the rPr of the first run that
overlaps the matched span. Tabs and other runs in the paragraph are left in
place — only the matched text changes. So this is the no-rebuild way to FILL a
formatted/tabbed template line: \`replace "Organization Name" "Northwind Robotics"\`
on a "**Organization Name**⇥Date" line keeps the bold and the tab; don't
hand-build \`edit --runs\` JSON to refill a line. \`--batch\` fills many at once.
When a single invocation produces multiple replacements in the same paragraph,
they're applied in reverse offset order so earlier offsets don't shift before
being applied.

By default the PATTERN is normalized: balanced markdown emphasis around
non-whitespace (**X**, __X__, *X*, \`X\`) is stripped; smart quotes match
straight quotes; em-dash and en-dash match the hyphen. The REPLACEMENT
is always literal — whatever bytes you pass go in as-is. Pass --exact
to match the raw pattern verbatim. --regex is always verbatim.

If the PATTERN or REPLACEMENT begins with a dash (a negative number, "-$500.00",
"--TODO"), put a bare "--" before the positionals so it isn't parsed as a flag:
  docx replace doc.docx -- "Total" "-$500.00"

With --regex, REPLACEMENT supports JS String.replace substitution syntax:
  $1, $2, ...   numbered capture groups
  $&            the matched substring
  $\`            text before the match
  $'            text after the match
  $$            a literal $

Output:
  Prints a one-line confirmation on success (exit 0) — replace mutates text in place and mints no new
  addressable handle (matched-span locators shift as text changes; re-read or
  use --dry-run to see them). --verbose / --dry-run print
  {ok:true, operation, totalMatches, replaced, matches:[{locator,…}], …}.
  Errors print {code, error, hint?} with a nonzero exit.

Examples:
  docx replace doc.docx "fox" "cat"
  docx replace doc.docx "fox" "cat" --all
  docx replace doc.docx "TODO|FIXME" "DONE" --regex --all
  docx replace doc.docx "(\\w+) (\\w+)" "$2 $1" --regex --all
  docx replace doc.docx "wordy phrase" "tighter phrase" --all --dry-run
  docx replace doc.docx --batch edits.jsonl

Batch JSONL example (one substitution per line, applied in order):
  {"pattern": "Q2", "replacement": "Q3", "all": true}
  {"pattern": "FY24", "replacement": "FY25"}
  {"pattern": "(\\\\w+)@old\\\\.com", "replacement": "$1@new.com", "regex": true, "all": true}
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			batch: { type: "string" },
			regex: { type: "boolean" },
			"ignore-case": { type: "boolean" },
			all: { type: "boolean" },
			limit: { type: "string" },
			author: { type: "string" },
			track: { type: "boolean" },
			current: { type: "boolean" },
			baseline: { type: "boolean" },
			exact: { type: "boolean" },
			...SAVE_FLAGS,
		},
		HELP,
	);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const batchInput = parsed.values.batch as string | undefined;
	if (batchInput !== undefined) {
		if (parsed.positionals.length > 1) {
			return fail(
				"USAGE",
				"--batch reads pattern/replacement from the JSONL file; don't also pass them as positionals",
				HELP,
			);
		}
		return runReplaceBatch(path, batchInput, parsed.values);
	}

	const pattern = parsed.positionals[1];
	const replacement = parsed.positionals[2];
	if (pattern == null) return fail("USAGE", "Missing PATTERN argument", HELP);
	if (replacement == null) {
		return fail("USAGE", "Missing REPLACEMENT argument", HELP);
	}

	const ignoreCase = Boolean(parsed.values["ignore-case"]);
	const useRegex = Boolean(parsed.values.regex);
	const wantAll = Boolean(parsed.values.all);
	const exact = Boolean(parsed.values.exact);
	const findView = resolveView(parsed.values);
	if (!findView) {
		return fail("USAGE", "--current and --baseline are mutually exclusive");
	}
	const limitRaw = parsed.values.limit as string | undefined;
	const limit = limitRaw === undefined ? undefined : Number(limitRaw);
	if (limit !== undefined && (!Number.isInteger(limit) || limit <= 0)) {
		return fail(
			"USAGE",
			`--limit must be a positive integer, got "${limitRaw}"`,
		);
	}

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	let findResult: ReturnType<typeof findTextSpans>;
	try {
		findResult = findTextSpans(document.body, pattern, {
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
		return fail("USAGE", `Invalid pattern: ${message}`);
	}

	const allMatches = findResult.matches;
	const normalizationFields =
		findResult.normalizedQuery !== undefined
			? {
					normalizedPattern: findResult.normalizedQuery,
					normalizationApplied: findResult.normalizationApplied,
				}
			: {};

	let selected: TextMatch[];
	if (limit !== undefined) {
		selected = allMatches.slice(0, limit);
	} else if (wantAll) {
		selected = allMatches;
	} else {
		selected = allMatches.slice(0, 1);
	}

	const matchesPayload = selected.map((match) => ({
		locator: `${match.blockId}:${match.start}-${match.end}`,
		blockId: match.blockId,
		start: match.start,
		end: match.end,
		text: match.text,
	}));

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			operation: "replace",
			dryRun: true,
			path,
			pattern,
			replacement,
			regex: useRegex,
			ignoreCase,
			view: findView,
			totalMatches: allMatches.length,
			replaced: selected.length,
			matches: matchesPayload,
			...normalizationFields,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	if (selected.length === 0) {
		await respondAck({
			ok: true,
			operation: "replace",
			path,
			pattern,
			replacement,
			regex: useRegex,
			ignoreCase,
			view: findView,
			totalMatches: 0,
			replaced: 0,
			matches: [],
			...normalizationFields,
		});
		return EXIT.OK;
	}

	// Apply in reverse document order so earlier offsets stay valid as later
	// ones get rewritten — both within a paragraph and across paragraphs.
	const reversed = [...selected].sort((leftMatch, rightMatch) => {
		if (leftMatch.blockId !== rightMatch.blockId) {
			return rightMatch.blockId.localeCompare(leftMatch.blockId);
		}
		return rightMatch.start - leftMatch.start;
	});

	const authorFlag = parsed.values.author as string | undefined;
	const tracked: TrackedReplaceOptions | undefined = resolveTracked(
		document,
		parsed.values.track,
	)
		? {
				meta: { author: resolveAuthor(authorFlag), date: resolveDate() },
				allocator: new TrackChanges(document).createAllocator(),
			}
		: undefined;

	const regexFlags = ignoreCase ? "i" : "";
	for (const match of reversed) {
		const concreteReplacement = useRegex
			? match.text.replace(new RegExp(pattern, regexFlags), replacement)
			: replacement;
		const blockRef = document.body.resolveBlock(match.blockId);
		replaceSpanInParagraph(
			blockRef.node,
			{ start: match.start, end: match.end },
			concreteReplacement,
			tracked,
			findView,
		);
	}

	await document.save(outputPath);

	// A default (first-match) or --limit replace can leave matches behind. Say so
	// in the text confirmation — a weak agent that ran a bare `replace` and saw
	// "1 occurrence replaced" otherwise assumes it got them all (the résumé agent
	// errored twice before discovering --all). Silent on a full sweep.
	const remaining = allMatches.length - selected.length;
	const partialHint =
		remaining > 0
			? `↳ ${remaining} more match${remaining === 1 ? "" : "es"} left unreplaced (${selected.length} of ${allMatches.length} done) — pass --all to replace every match, or --limit N for a specific count.`
			: undefined;

	await respondAck(
		{
			ok: true,
			operation: "replace",
			path: outputPath ?? path,
			pattern,
			replacement,
			regex: useRegex,
			ignoreCase,
			view: findView,
			totalMatches: allMatches.length,
			replaced: selected.length,
			matches: matchesPayload,
			...normalizationFields,
		},
		partialHint,
	);
	return EXIT.OK;
}
