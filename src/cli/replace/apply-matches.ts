import type { Body } from "@core";
import { EquationParseError } from "@core/equation";
import {
	equationSpanRefusal,
	type ReplacementFormatting,
	replaceMatches,
	StaleEquationError,
	type TextMatch,
	type TrackedReplaceOptions,
} from "@core/find";
import { spanLocator } from "../parse-helpers";
import { fail } from "../respond";

/** Refuse the selected matches when any meets an equation in a way replace
 *  can't carry out (`equationSpanRefusal`: a cut token, or a shell-mangled
 *  replacement that would delete one), else apply them — whole tokens
 *  deleted, inside-token spans recompiled — and return undefined. Shared by
 *  single-shot and `--batch` replace; `entry` prefixes a batch entry's error.
 *  Every refusal comes before any mutation, and the caller applies even on
 *  `--dry-run` (skipping only the save), so a preview refuses exactly what the
 *  real replace would — LaTeX that doesn't parse included. */
export async function applyMatchesOrFail(
	body: Body,
	selected: TextMatch[],
	expand: (matchText: string) => string,
	options: {
		tracked?: TrackedReplaceOptions | undefined;
		view: TextMatchView;
		formatting?: ReplacementFormatting | undefined;
		entry?: number;
	},
): Promise<number | undefined> {
	const prefix = options.entry === undefined ? "" : `entry ${options.entry}: `;
	const replacements = selected.map((match) => ({
		blockId: match.blockId,
		start: match.start,
		end: match.end,
		replacement: expand(match.text),
		equations: match.equations,
		matched: match.text,
	}));
	for (const replacement of replacements) {
		const refusal = equationSpanRefusal(
			spanLocator(replacement),
			replacement.equations,
			replacement.replacement,
			"replace",
			{ pattern: replacement.matched },
		);
		if (refusal) {
			return fail(refusal.code, `${prefix}${refusal.message}`, refusal.hint);
		}
	}
	try {
		replaceMatches(body, replacements, options);
	} catch (error) {
		if (error instanceof EquationParseError) {
			return fail(
				"USAGE",
				`${prefix}the replacement leaves LaTeX that doesn't parse (${JSON.stringify(error.latex)}): ${error.message}. Nothing was changed.`,
				"A match inside an equation rewrites its LaTeX, so the result must still be valid LaTeX. Check the replacement, or set the whole equation with `docx equations edit FILE --at eqN --equation '…'`.",
			);
		}
		if (error instanceof StaleEquationError) {
			return fail(
				"UNSUPPORTED",
				`${prefix}${error.message}`,
				`Match text outside it, or set the whole equation with \`docx equations edit FILE --at ${error.equationId} --equation '…'\`.`,
			);
		}
		throw error;
	}
	return undefined;
}

type TextMatchView = Parameters<typeof replaceMatches>[2]["view"];
