import { resolveAuthor, resolveDate, TrackChanges } from "@core";
import {
	findTextSpans,
	replacementExpander,
	replaceSpanInParagraph,
	selectMatches,
	type TrackedReplaceOptions,
} from "@core/find";
import {
	decodeInlineEscapes,
	resolveView,
	spanLocator,
} from "../parse-helpers";
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
import { partialReplaceHint, runReplaceAcross } from "./across";
import { runReplaceBatch } from "./batch";
import { matchesInScope, resolveReplaceScope, ScopeError } from "./scope";

const HELP = `docx replace — substitute content (sed for docx)

Usage:
  docx replace FILE PATTERN REPLACEMENT [options]
  docx replace FILE --batch FILE.jsonl [options]   # a sed-script, one read
  docx replace FILE --batch -          [options]   # read JSONL from stdin

Examples:
  # Replace many placeholders/terms in ONE call — the preferred path (a sed-script).
  # Write one substitution (JSON object) per line to a file, then apply it:
  #   subs.jsonl:
  #     {"pattern":"Q2","replacement":"Q3","all":true}
  #     {"pattern":"FY24","replacement":"FY25"}
  #     {"pattern":"(\\\\w+)@old\\\\.com","replacement":"$1@new.com","regex":true,"all":true}
  #     {"pattern":"City, State","replacement":"Boston, MA","at":"p20"}
  docx replace doc.docx --batch subs.jsonl
  # …or one at a time:
  docx replace doc.docx "fox" "cat" --all
  docx replace doc.docx "TODO|FIXME" "DONE" --regex --all
  docx replace doc.docx "(\\w+) (\\w+)" "$2 $1" --regex --all
  docx replace doc.docx "wordy phrase" "tighter phrase" --all --dry-run

Options:
  --at LOCATOR      confine the replace to ONE paragraph (a body pN or a cell
                    paragraph tT:rRcC:pN). Use when the same placeholder repeats
                    across the document and you want the one in a specific
                    paragraph:
                      docx replace doc.docx --at p20 "City, State" "Boston, MA"
  --regex           treat PATTERN as a JavaScript regular expression
  --ignore-case     case-insensitive match
  --all             replace every match (default: just the first; with --at,
                    scoped to that paragraph)
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

Batch (--batch PATH | -):
  A sed-script from one read — the preferred way to run several substitutions.
  Each JSONL line is one substitution whose keys mirror the flags:
  {"pattern": …, "replacement": …} plus optional "regex"/"ignoreCase"/"all"/
  "limit"/"exact"/"at". Entries apply in order, each seeing the previous
  entry's edits (like running replace repeatedly). Don't pass the PATTERN/
  REPLACEMENT positionals or --at alongside --batch.

Matching:
  By default the PATTERN is normalized so text copied from \`docx read\`
  matches the real document text: markdown emphasis is stripped (**X**
  matches X), smart quotes match straight quotes, em/en dashes match "-".
  The REPLACEMENT is always literal — whatever you pass goes in as-is, and it
  inherits the formatting of the matched span (bold text stays bold).
  --exact matches the raw pattern verbatim; --regex is always verbatim.
  If PATTERN or REPLACEMENT begins with a dash ("-$500.00", "--TODO"), put a
  bare "--" before the positionals: docx replace doc.docx -- "Total" "-$500.00"
  With --regex, REPLACEMENT supports $1, $2, … (capture groups), $& (the
  whole match), and $$ (a literal $).

Multi-line (editor-style):
  \\n and \\t in PATTERN/REPLACEMENT are decoded to real characters. A "\\n"
  in the PATTERN matches an in-paragraph line break OR the boundary between
  consecutive paragraphs (in the body or within one table cell — never across
  a cell wall). The REPLACEMENT's "\\n"s then define the resulting structure,
  as if the span were selected in Word and the replacement typed:
    docx replace doc.docx "One.\\nTwo." "One. Two."   # MERGES two paragraphs
    docx replace doc.docx "One. Two." "One.\\nTwo."   # SPLITS a paragraph
  Untracked only — with tracking on (or --track) it refuses; turn tracking
  off first. Block ids shift afterward; re-read before more edits.

Output:
  Prints a one-line confirmation with the replaced count on success (exit 0).
  --verbose / --dry-run print {ok:true, operation, totalMatches, replaced,
  matches:[{locator,…}], …}. Errors print {code, error, hint?} + nonzero exit.
  A PATTERN that matches NOTHING is an error, not a silent success — 0
  occurrences exits nonzero (MATCH_NOT_FOUND). Unsure it matches? Probe with
  \`docx find FILE PATTERN\` or --dry-run and READ the reported count — both
  exit 0 whether or not it matches; only the real replace exits nonzero on 0.
  --batch: if ANY entry matches nothing, the exit is nonzero — but the entries
  that DID match are already applied and SAVED. Nonzero means "some entries
  missed," not "nothing changed" (the error names the misses).
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			batch: { type: "string" },
			at: { type: "string" },
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
		if (parsed.values.at !== undefined) {
			return fail(
				"USAGE",
				'--at is per-entry in batch mode: put "at" on each JSONL line, not on the command',
				HELP,
			);
		}
		return runReplaceBatch(path, batchInput, parsed.values);
	}

	// Inline argv decode, same as every authoring surface: `\n`/`\t` typed in a
	// double-quoted shell argument become real characters. A real "\n" is also
	// the cross-paragraph gate below.
	const pattern = decodeInlineEscapes(parsed.positionals[1]);
	const replacement = decodeInlineEscapes(parsed.positionals[2]);
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

	// A "\n" in the pattern spans lines (line break or paragraph boundary); a
	// "\n" in the replacement inserts a paragraph mark. Either routes to the
	// cross-paragraph path — editor-style, untracked only (it refuses under
	// tracking rather than record a merge/split wrong).
	if (pattern.includes("\n") || replacement.includes("\n")) {
		if (parsed.values.at !== undefined) {
			return fail(
				"USAGE",
				"--at scopes a replace to ONE paragraph, but a \\n-bearing pattern/replacement spans paragraph boundaries — drop --at, or remove the \\n",
			);
		}
		return runReplaceAcross(path, pattern, replacement, {
			regex: useRegex,
			ignoreCase,
			exact,
			all: wantAll,
			...(limit !== undefined ? { limit } : {}),
			view: findView,
			track: Boolean(parsed.values.track),
			...(parsed.values.output !== undefined
				? { output: parsed.values.output as string }
				: {}),
			dryRun: Boolean(parsed.values["dry-run"]),
		});
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

	// `--at pN` scopes the substitution to one paragraph (see ./scope) — the
	// résumé fix for repeated placeholders. Applied to the match set BEFORE
	// first/all/limit selection, so "first match" means first WITHIN the scope.
	const atScope = parsed.values.at as string | undefined;
	let allMatches = findResult.matches;
	if (atScope !== undefined) {
		try {
			const blockId = resolveReplaceScope(document, atScope);
			allMatches = matchesInScope(allMatches, blockId);
		} catch (scopeError) {
			if (scopeError instanceof ScopeError) {
				return fail(scopeError.code, scopeError.message);
			}
			throw scopeError;
		}
	}
	const normalizationFields =
		findResult.normalizedQuery !== undefined
			? {
					normalizedPattern: findResult.normalizedQuery,
					normalizationApplied: findResult.normalizationApplied,
				}
			: {};

	const selected = selectMatches(allMatches, { all: wantAll, limit });

	const matchesPayload = selected.map((match) => ({
		locator: spanLocator(match),
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
			...(atScope ? { at: atScope } : {}),
			totalMatches: allMatches.length,
			replaced: selected.length,
			matches: matchesPayload,
			...normalizationFields,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	// A zero-match replace is a SILENT NO-OP — the document is unchanged. Exit
	// nonzero (MATCH_NOT_FOUND) instead of a cheerful "0 occurrences replaced":
	// weak agents key their done/retry decision off the exit code (they demonstrably
	// react to nonzero and ignore a 0-count success line), so a 0 exit here bakes in
	// a confidently-wrong document. Nonzero forces them to notice and fix.
	if (selected.length === 0) {
		const scopeNote = atScope ? ` within ${atScope}` : "";
		return await fail(
			"MATCH_NOT_FOUND",
			`Pattern not found${scopeNote}: ${JSON.stringify(pattern)} — 0 occurrences, nothing changed.`,
			`Match LITERAL document text, not read-view markup (\`<mark>\`/\`<u>\`/\`**\`…) or a locator. Run \`docx find ${path} ${JSON.stringify(pattern)}\` to see if/where it occurs; add --ignore-case, --regex, --at <locator> to scope, or --current/--baseline to search tracked-change text.`,
		);
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

	const expand = replacementExpander({
		pattern,
		replacement,
		regex: useRegex,
		ignoreCase,
	});
	for (const match of reversed) {
		const blockRef = document.body.resolveBlock(match.blockId);
		replaceSpanInParagraph(
			blockRef.node,
			{ start: match.start, end: match.end },
			expand(match.text),
			tracked,
			findView,
		);
	}

	await document.save(outputPath);

	const partialHint = partialReplaceHint(selected.length, allMatches.length);

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
			...(atScope ? { at: atScope } : {}),
			totalMatches: allMatches.length,
			replaced: selected.length,
			matches: matchesPayload,
			...normalizationFields,
		},
		partialHint,
	);
	return EXIT.OK;
}
