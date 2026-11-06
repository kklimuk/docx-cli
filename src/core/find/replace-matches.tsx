import type { Body } from "../ast/document/body";
import { latexToOmml } from "../equation/emit";
import {
	batchFileHint,
	shellMangledEquationLoss,
} from "../equation/shell-mangle";
import { formatEquationToken } from "../equation/token";
import { removeFieldsCoveredInParagraph } from "../fields/remove-covered-fields";
import type { XmlNode } from "../parser";
import { Del, Ins } from "../track-changes/emit";
import type { EquationOverlap, FindView } from "./index";
import {
	type ReplacementFormatting,
	replaceSpanInParagraph,
	type TrackedReplaceOptions,
} from "./replace-span";

/** Apply span replacements — `replace`'s matches, or one `edit --at pN:S-E` —
 *  with equations as tokens in the text. Per paragraph, in reverse offset
 *  order so every earlier span's offsets stay valid:
 *
 *  - a span with no equation, or one COVERING whole equation tokens, is an
 *    ordinary span replace (`replaceSpanInParagraph`), which deletes each
 *    covered equation with the text around it (`<w:del>` when tracked);
 *  - a span strictly INSIDE one token rewrites that equation's LaTeX — every
 *    such span in one equation folds into ONE recompile (`latexToOmml`), as
 *    `equations edit` does it: in place untracked, a `<w:del>OLD</w:del>
 *    <w:ins>NEW</w:ins>` pair tracked. A rewrite that leaves no LaTeX deletes
 *    the equation.
 *
 *  A span that CUTS a token must have been refused (`equationSpanRefusal`)
 *  first. The replacement text is always literal: a `$5` in it is text. */
export function replaceMatches(
	body: Body,
	matches: SpanReplacement[],
	options: ReplaceOptions,
): void {
	const byBlock = new Map<string, SpanReplacement[]>();
	for (const match of matches) {
		const group = byBlock.get(match.blockId) ?? [];
		group.push(match);
		byBlock.set(match.blockId, group);
	}
	for (const [blockId, group] of byBlock) {
		replaceInParagraph(
			body.resolveBlock(blockId).node,
			group,
			(id) => body.equationReferences.get(id)?.node,
			options,
		);
	}
}

/** A span to replace, with the equations it touches (`equationOverlaps`). */
export type SpanReplacement = {
	blockId: string;
	start: number;
	end: number;
	replacement: string;
	equations?: EquationOverlap[] | undefined;
};

export type ReplaceOptions = {
	tracked?: TrackedReplaceOptions | undefined;
	view: FindView;
	formatting?: ReplacementFormatting | undefined;
};

/** `replaceMatches` for one paragraph whose node the caller already holds
 *  (`Edit.span`). `equationNode` resolves an `eqN` to its live `<m:oMath>`. */
export function replaceInParagraph(
	paragraph: XmlNode,
	matches: SpanReplacement[],
	equationNode: (id: string) => XmlNode | undefined,
	{ tracked, view, formatting }: ReplaceOptions,
): void {
	const operations: { start: number; apply: () => void }[] = [];
	const rewrites = new Map<
		string,
		{ overlap: EquationOverlap; changes: SpanReplacement[] }
	>();
	for (const match of matches) {
		const inside = match.equations?.find(
			(overlap) => overlap.relation === "inside",
		);
		if (inside) {
			const rewrite = rewrites.get(inside.id) ?? {
				overlap: inside,
				changes: [],
			};
			rewrite.changes.push(match);
			rewrites.set(inside.id, rewrite);
			continue;
		}
		operations.push({
			start: match.start,
			apply: () => {
				const span = { start: match.start, end: match.end };
				// A field the span covers goes whole — its markers first, then
				// the walker takes its result with the rest of the text.
				removeFieldsCoveredInParagraph(
					paragraph,
					span,
					view,
					tracked && (() => mintMeta(tracked)),
				);
				replaceSpanInParagraph(
					paragraph,
					span,
					match.replacement,
					tracked,
					view,
					formatting,
				);
			},
		});
	}
	for (const { overlap, changes } of rewrites.values()) {
		const node = equationNode(overlap.id);
		if (!node) throw new Error(`replaceInParagraph: ${overlap.id} not found`);
		operations.push({
			start: overlap.start,
			apply: () =>
				rewriteEquation(
					paragraph,
					node,
					rewrittenLatex(overlap, changes),
					overlap.display,
					tracked,
				),
		});
	}
	operations.sort((left, right) => right.start - left.start);
	for (const operation of operations) operation.apply();
}

/** The equation's LaTeX with each inside-span replaced, right to left. */
function rewrittenLatex(
	overlap: EquationOverlap,
	changes: SpanReplacement[],
): string {
	const contentStart = overlap.start + (overlap.display ? 2 : 1);
	let latex = overlap.latex;
	for (const change of [...changes].sort((a, b) => b.start - a.start)) {
		const from = change.start - contentStart;
		const to = change.end - contentStart;
		latex = latex.slice(0, from) + change.replacement + latex.slice(to);
	}
	return latex;
}

/** Recompile `node` from `latex` where it stands in the live paragraph —
 *  located by identity, never through a cached parent array, which an earlier
 *  span replace in the same paragraph may have swapped out. Throws
 *  `EquationParseError` for LaTeX temml rejects. */
function rewriteEquation(
	paragraph: XmlNode,
	node: XmlNode,
	latex: string,
	display: boolean,
	tracked: TrackedReplaceOptions | undefined,
): void {
	const siblings = liveSiblings(paragraph, node);
	if (!siblings) throw new Error("rewriteEquation: equation not in paragraph");
	const index = siblings.indexOf(node);
	const replacement =
		latex.trim().length > 0 ? latexToOmml(latex, display) : null;
	if (!tracked) {
		siblings.splice(index, 1, ...(replacement ? [replacement] : []));
		return;
	}
	const deletion = <Del meta={mintMeta(tracked)}>{[node]}</Del>;
	siblings.splice(
		index,
		1,
		deletion,
		...(replacement
			? [<Ins meta={mintMeta(tracked)}>{[replacement]}</Ins>]
			: []),
	);
}

function liveSiblings(container: XmlNode, node: XmlNode): XmlNode[] | null {
	if (container.children.includes(node)) return container.children;
	for (const child of container.children) {
		if (child.isText) continue;
		const found = liveSiblings(child, node);
		if (found) return found;
	}
	return null;
}

function mintMeta(tracked: TrackedReplaceOptions) {
	return { ...tracked.meta, revisionId: tracked.allocator.next() };
}

/** Why a span replace touching equations can't go ahead, or null. Shared by
 *  every replace surface and `edit --at pN:S-E`, so they refuse alike:
 *
 *  - UNSUPPORTED when the span CUTS a token (crosses a `$` delimiter — neither
 *    the whole equation nor text inside it), or puts a paragraph break inside
 *    an equation (`multiline`: a cross-paragraph replace);
 *  - USAGE when the span deletes equations and the replacement carries a
 *    shell-mangling telltale (`shellMangledEquationLoss`).
 *
 *  `command` names the surface for the batch-file hint. */
export function equationSpanRefusal(
	locator: string,
	overlaps: EquationOverlap[] | undefined,
	replacement: string,
	command: "replace" | "edit",
	{
		multiline = false,
		pattern,
	}: {
		multiline?: boolean;
		/** The text the span matched — the batch-hint example's pattern. */
		pattern?: string;
	} = {},
): EquationRefusal | null {
	if (!overlaps || overlaps.length === 0) return null;
	const cut = overlaps.find((overlap) => overlap.relation === "cut");
	if (cut) {
		const token = tokenOf(cut);
		// A pattern bash gutted (`"$E=mc^2$ in"` → `=mc^2$ in`) matches from
		// inside the token out past its closer: show the line it meant.
		const restored =
			pattern === undefined
				? undefined
				: shellMangledEquationLoss(pattern, [cut])?.repaired;
		const shellNote = restored
			? ` The pattern looks like the shell expanded its "$" away: ${surfaceBatchHint(command, cut, { pattern: restored, replacement })}`
			: "";
		return {
			code: "UNSUPPORTED",
			message: `${locator}: the span cuts ${cut.id} (${token}) — it crosses the equation's ${cut.display ? "$$" : "$"} delimiter, so it is neither the whole equation nor text inside it. Nothing was changed.`,
			hint:
				command === "replace"
					? `In find/replace text an equation is its token exactly as \`docx read\` shows it. Match the whole token (${token}) to delete the equation, only text inside it (e.g. ${cut.latex}) to rewrite its LaTeX, or only the text beside it.${shellNote}`
					: `In span offsets an equation is its token exactly as \`docx read\` shows it (${token} is ${cut.end - cut.start} characters, ${cut.blockId}:${cut.start}-${cut.end}). Cover the whole token to delete the equation, stay strictly inside it to rewrite its LaTeX, or stay beside it — \`docx find FILE "TEXT"\` gives exact offsets.`,
		};
	}
	const inside = overlaps.find((overlap) => overlap.relation === "inside");
	if (inside && (multiline || replacement.includes("\n"))) {
		return {
			code: "UNSUPPORTED",
			message: `${locator}: the replacement would put a paragraph break inside ${inside.id} (${tokenOf(inside)}). Nothing was changed.`,
			hint: `Rewrite the equation's LaTeX without the newline (\`docx equations edit FILE --at ${inside.id}\`), or match text outside it.`,
		};
	}
	const covered = overlaps.filter((overlap) => overlap.relation === "covered");
	const mangled = shellMangledEquationLoss(replacement, covered);
	if (mangled) {
		const [first] = covered;
		return {
			code: "USAGE",
			message: `${locator}: the replacement deletes ${covered.map((overlap) => `${overlap.id} (${tokenOf(overlap)})`).join(", ")}, and ${mangled.reason}. Nothing was changed.`,
			hint: first
				? `${surfaceBatchHint(command, first, { pattern, replacement: mangled.repaired ?? replacement })} To delete the equation on purpose, replace its whole token with plain text that has no stray $.`
				: "",
		};
	}
	return null;
}

export type EquationRefusal = {
	code: "UNSUPPORTED" | "USAGE";
	message: string;
	hint: string;
};

function tokenOf(overlap: EquationOverlap): string {
	return formatEquationToken(overlap.latex, "", overlap.display);
}

/** The batch-file hint with an example line for this surface: the line the
 *  agent meant to send, as far as it can be told — the matched text and the
 *  (repaired) replacement for `replace`, the token in context for `edit`. */
function surfaceBatchHint(
	command: "replace" | "edit",
	overlap: EquationOverlap,
	{ pattern, replacement }: { pattern?: string; replacement: string },
): string {
	const token = tokenOf(overlap);
	if (command === "replace") {
		return batchFileHint("replace", "subs.jsonl", {
			pattern: pattern ?? token,
			replacement,
		});
	}
	return batchFileHint("edit", "edits.jsonl", {
		at: overlap.blockId,
		markdown: `… ${token} …`,
	});
}
