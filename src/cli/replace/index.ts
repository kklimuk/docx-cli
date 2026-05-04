import { resolveBlock, saveDocView } from "@core";
import { findTextSpans, type TextMatch } from "@core/find";
import { parseArgs } from "util";
import { EXIT, fail, openOrFail, respond, writeStdout } from "../respond";
import {
	replaceSpanInParagraph,
	TrackedChangeBoundaryError,
} from "./replace-span";

const HELP = `docx replace — substitute text spans (sed for docx)

Usage:
  docx replace FILE PATTERN REPLACEMENT [options]

Options:
  --regex           treat PATTERN as a JavaScript regular expression
  --ignore-case     case-insensitive match
  --all             replace every match (default: just the first)
  --limit N         replace at most N matches (in document order)
  -o, --output PATH write to PATH instead of overwriting FILE
  --dry-run         report what would change without writing the file
  -h, --help        show this help

Within-paragraph matches only. Run formatting (rPr) on the surrounding text
is preserved; the replacement run inherits the rPr of the first run that
overlaps the matched span. When a single invocation produces multiple
replacements in the same paragraph, they're applied in reverse offset order
so earlier offsets don't shift before being applied.

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
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				regex: { type: "boolean" },
				"ignore-case": { type: "boolean" },
				all: { type: "boolean" },
				limit: { type: "string" },
				output: { type: "string", short: "o" },
				"dry-run": { type: "boolean" },
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
	const limitRaw = parsed.values.limit as string | undefined;
	const limit = limitRaw === undefined ? undefined : Number(limitRaw);
	if (limit !== undefined && (!Number.isInteger(limit) || limit <= 0)) {
		return fail(
			"USAGE",
			`--limit must be a positive integer, got "${limitRaw}"`,
		);
	}

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	let allMatches: TextMatch[];
	try {
		allMatches = findTextSpans(view.doc, pattern, {
			regex: useRegex,
			ignoreCase,
		});
	} catch (matcherError) {
		const message =
			matcherError instanceof Error
				? matcherError.message
				: String(matcherError);
		return fail("USAGE", `Invalid pattern: ${message}`);
	}

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
			totalMatches: allMatches.length,
			replaced: selected.length,
			matches: matchesPayload,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	if (selected.length === 0) {
		await respond({
			ok: true,
			operation: "replace",
			path,
			pattern,
			replacement,
			regex: useRegex,
			ignoreCase,
			totalMatches: 0,
			replaced: 0,
			matches: [],
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

	const regexFlags = ignoreCase ? "i" : "";
	try {
		for (const match of reversed) {
			const concreteReplacement = useRegex
				? match.text.replace(new RegExp(pattern, regexFlags), replacement)
				: replacement;
			const blockRef = resolveBlock(view, match.blockId);
			replaceSpanInParagraph(
				blockRef.node,
				{ start: match.start, end: match.end },
				concreteReplacement,
			);
		}
	} catch (error) {
		if (error instanceof TrackedChangeBoundaryError) {
			return fail(
				"TRACKED_CHANGE_CONFLICT",
				error.message,
				"Use `docx track-changes off` (or accept/reject the change in Word) before replacing.",
			);
		}
		throw error;
	}

	await saveDocView(view, outputPath);

	await respond({
		ok: true,
		operation: "replace",
		path: outputPath ?? path,
		pattern,
		replacement,
		regex: useRegex,
		ignoreCase,
		totalMatches: allMatches.length,
		replaced: selected.length,
		matches: matchesPayload,
	});
	return EXIT.OK;
}
