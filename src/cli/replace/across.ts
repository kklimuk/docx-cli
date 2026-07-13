import {
	type AcrossReplaceResult,
	applyAcrossReplace,
	type FindView,
} from "@core/find";
import { spanLocator } from "../parse-helpers";
import {
	EXIT,
	fail,
	openOrFail,
	resolveTracked,
	respond,
	respondAck,
} from "../respond";

/** `docx replace` with a "\n" in the PATTERN or REPLACEMENT — the
 *  cross-paragraph path. Each "\n" in the pattern matches an in-paragraph line
 *  break OR a paragraph boundary (consecutive paragraphs in the body or within
 *  one table cell); the replacement's newlines then define the resulting
 *  paragraph structure, editor-style — replacing across a boundary with a
 *  single-line replacement MERGES the paragraphs, and a "\n" in the
 *  replacement inserts a paragraph mark (splitting, when the match sat inside
 *  one paragraph).
 *
 *  Untracked only: OOXML records a merge/split on the paragraph MARK
 *  (`<w:pPr><w:rPr><w:ins|del>`), which this path doesn't emit yet — so under
 *  tracking it refuses loudly instead of recording the change wrong or
 *  silently skipping the journal. */
export async function runReplaceAcross(
	path: string,
	pattern: string,
	replacement: string,
	options: {
		regex: boolean;
		ignoreCase: boolean;
		exact: boolean;
		all: boolean;
		limit?: number;
		view: FindView;
		track: boolean;
		output?: string;
		dryRun: boolean;
	},
): Promise<number> {
	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	if (resolveTracked(document, options.track)) {
		return fail("USAGE", ACROSS_TRACKED_MESSAGE, ACROSS_TRACKED_HINT);
	}

	let result: AcrossReplaceResult;
	try {
		result = applyAcrossReplace(
			document.body,
			{
				pattern,
				replacement,
				regex: options.regex,
				ignoreCase: options.ignoreCase,
				exact: options.exact,
				all: options.all,
				...(options.limit !== undefined ? { limit: options.limit } : {}),
				view: options.view,
			},
			{ dryRun: options.dryRun },
		);
	} catch (matcherError) {
		const message =
			matcherError instanceof Error
				? matcherError.message
				: String(matcherError);
		return fail("USAGE", `Invalid pattern: ${message}`);
	}

	const matchesPayload = result.replaced.map((match) => ({
		locator: spanLocator(match),
		text: match.text,
	}));
	const normalizationFields =
		result.normalizedQuery !== undefined
			? {
					normalizedPattern: result.normalizedQuery,
					normalizationApplied: result.normalizationApplied,
				}
			: {};

	if (options.dryRun) {
		await respond({
			operation: "replace",
			dryRun: true,
			path,
			pattern,
			replacement,
			regex: options.regex,
			ignoreCase: options.ignoreCase,
			view: options.view,
			crossParagraph: true,
			totalMatches: result.totalMatches,
			replaced: result.replaced.length,
			matches: matchesPayload,
			...normalizationFields,
			...(options.output ? { output: options.output } : {}),
		});
		return EXIT.OK;
	}

	if (result.replaced.length === 0) {
		return await fail(
			"MATCH_NOT_FOUND",
			`Pattern not found: ${JSON.stringify(pattern)} — 0 occurrences, nothing changed.`,
			`Each \\n in the pattern matches a line break or the boundary between CONSECUTIVE paragraphs (never across a table, section break, or cell wall). \`docx read ${path}\` shows the lines — check the exact text on each side of every \\n.`,
		);
	}

	await document.save(options.output);

	const hints: string[] = [
		"↳ paragraph structure changed — block ids shifted; re-read before locator-addressed edits.",
	];
	const partialHint = partialReplaceHint(
		result.replaced.length,
		result.totalMatches,
	);
	if (partialHint) hints.push(partialHint);
	await respondAck(
		{
			ok: true,
			operation: "replace",
			path: options.output ?? path,
			pattern,
			replacement,
			regex: options.regex,
			ignoreCase: options.ignoreCase,
			view: options.view,
			crossParagraph: true,
			totalMatches: result.totalMatches,
			replaced: result.replaced.length,
			matches: matchesPayload,
			...normalizationFields,
		},
		hints.join("\n"),
	);
	return EXIT.OK;
}

/** Why a \n-bearing replace refuses under tracking — one wording shared by the
 *  single-shot and batch surfaces, so both entry points teach the same rule. */
export const ACROSS_TRACKED_MESSAGE =
	"a replace that crosses or inserts paragraph boundaries can't be recorded as a tracked change yet (OOXML tracks a merge/split on the paragraph mark, which this path doesn't emit) — it will NOT run silently untracked";

export const ACROSS_TRACKED_HINT =
	"Turn tracking off (docx track-changes off FILE) and re-run, or make within-paragraph replaces (no \\n) which track fine.";

/** The "N more matches left" nudge, shared by the in-paragraph and
 *  cross-paragraph paths: a weak agent that saw "1 occurrence replaced"
 *  assumes it got them all (the résumé agent errored twice before discovering
 *  --all), so a partial sweep says what's left and names the cure. Silent on a
 *  full sweep. */
export function partialReplaceHint(
	replaced: number,
	total: number,
): string | undefined {
	const remaining = total - replaced;
	if (remaining <= 0) return undefined;
	return `↳ ${remaining} more match${remaining === 1 ? "" : "es"} left unreplaced (${replaced} of ${total} done) — pass --all to replace every match, or --limit N for a specific count.`;
}
