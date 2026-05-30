import { resolveAuthor, resolveDate, TrackChanges } from "@core";
import {
	type FindView,
	findTextSpans,
	replaceSpanInParagraph,
	type TextMatch,
	type TrackedReplaceOptions,
} from "@core/find";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx replace — substitute text spans (sed for docx)

Usage:
  docx replace FILE PATTERN REPLACEMENT [options]

Options:
  --regex           treat PATTERN as a JavaScript regular expression
  --ignore-case     case-insensitive match
  --all             replace every match (default: just the first)
  --limit N         replace at most N matches (in document order)
  --author NAME     author for tracked changes (default: $DOCX_AUTHOR)
  --current         operate on the raw concatenation (both ins and del text)
  --baseline        operate on the pre-change text (skip ins/moveTo)
                    default: accepted document (skip del/moveFrom) — matches
                    "docx find" / "docx read"
  --exact           disable pattern normalization (no markdown-emphasis stripping,
                    no smart/straight quote or em/en-dash equivalence)
  -o, --output PATH write to PATH instead of overwriting FILE
  --dry-run         report what would change without writing the file
  -v, --verbose     print the success ack JSON (default: silent on success)
  -h, --help        show this help

Within-paragraph matches only. Run formatting (rPr) on the surrounding text
is preserved; the replacement run inherits the rPr of the first run that
overlaps the matched span. When a single invocation produces multiple
replacements in the same paragraph, they're applied in reverse offset order
so earlier offsets don't shift before being applied.

By default the PATTERN is normalized: balanced markdown emphasis around
non-whitespace (**X**, __X__, *X*, \`X\`) is stripped; smart quotes match
straight quotes; em-dash and en-dash match the hyphen. The REPLACEMENT
is always literal — whatever bytes you pass go in as-is. Pass --exact
to match the raw pattern verbatim. --regex is always verbatim.

With --regex, REPLACEMENT supports JS String.replace substitution syntax:
  $1, $2, ...   numbered capture groups
  $&            the matched substring
  $\`            text before the match
  $'            text after the match
  $$            a literal $

Examples:
  docx replace doc.docx "fox" "cat"
  docx replace doc.docx "fox" "cat" --all
  docx replace doc.docx "TODO|FIXME" "DONE" --regex --all
  docx replace doc.docx "(\\w+) (\\w+)" "$2 $1" --regex --all
  docx replace doc.docx "wordy phrase" "tighter phrase" --all --dry-run
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			regex: { type: "boolean" },
			"ignore-case": { type: "boolean" },
			all: { type: "boolean" },
			limit: { type: "string" },
			author: { type: "string" },
			current: { type: "boolean" },
			baseline: { type: "boolean" },
			exact: { type: "boolean" },
			output: { type: "string", short: "o" },
			"dry-run": { type: "boolean" },
			verbose: { type: "boolean", short: "v" },
			help: { type: "boolean", short: "h" },
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
	const pattern = parsed.positionals[1];
	const replacement = parsed.positionals[2];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);
	if (pattern == null) return fail("USAGE", "Missing PATTERN argument", HELP);
	if (replacement == null) {
		return fail("USAGE", "Missing REPLACEMENT argument", HELP);
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
			ok: true,
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
	const tracked: TrackedReplaceOptions | undefined =
		document.isTrackChangesEnabled()
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

	await respondAck({
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
	});
	return EXIT.OK;
}
