import { equationToken } from "../equation/token";
import {
	isEquation,
	isRunBearingWrapper,
	wrapperContent,
	type XmlNode,
} from "../parser";

/** How a whole-paragraph `--text` rewrite treats the paragraph's equations.
 *  The text is diffed against the paragraph's TOKEN view — what `read` shows,
 *  each visible equation as its `$latex$` / `$$latex$$` token — and each
 *  equation is decided by what the new text did with its token:
 *
 *  1. KEPT: the token appears verbatim (old equations matched in order). The
 *     equation object stays, untouched, wherever the diff puts it.
 *  2. EDITED: between the same kept neighbors the new text holds exactly as
 *     many well-formed `$…$` / `$$…$$` spans as there are missing equations,
 *     of the same kinds in the same order — the delimiters survived in place,
 *     only the LaTeX changed. Each is recompiled from its new LaTeX.
 *  3. DELETED: no such span is left between those neighbors.
 *  4. Everything else is literal text — a `$5` never becomes an equation.
 *
 *  Two shapes are ambiguous and refused (`ambiguity` is set): spans that
 *  don't pair one-to-one with the missing equations (two spans for one
 *  equation, an inline span where a display block was), and a missing
 *  equation's token reappearing somewhere else (a reordering — `--text` has no
 *  way to say "move this object", and writing it as literal text beside a
 *  deleted equation is never what was meant).
 *
 *  The returned `text` is the new text with every kept token and edited span
 *  replaced by a private-use SENTINEL (`keepSentinel(i)` / `editSentinel(i)`,
 *  `i` the equation's index in `visibleEquations` order) that `tokenize`
 *  keeps as one diff token, so the word diff aligns a kept equation exactly
 *  as it aligns a kept word. */
export function planEquationText(
	paragraph: XmlNode,
	newText: string,
): EquationTextPlan {
	const equations = visibleEquations(paragraph);
	// Private-use code points are the sentinel alphabet; nothing a human or an
	// agent types contains them, but strip any so a stray one can't alias.
	const text = newText.replace(SENTINEL_CHARS, "");
	if (equations.length === 0) {
		return { equations, text, edited: [], deleted: [] };
	}

	const anchors = keptAnchors(equations, text);
	const edited: EquationTextPlan["edited"] = [];
	const deleted: OldEquation[] = [];
	const replacements: { start: number; end: number; sentinel: string }[] = [];
	for (const anchor of anchors) {
		replacements.push({
			start: anchor.position,
			end: anchor.position + anchor.equation.token.length,
			sentinel: keepSentinel(anchor.equation.index),
		});
	}

	// The gaps between consecutive kept anchors, in old and new text alike.
	const bounds = [
		{ oldIndex: -1, newEnd: 0 },
		...anchors.map((anchor) => ({
			oldIndex: anchor.equation.index,
			newEnd: anchor.position + anchor.equation.token.length,
			newStart: anchor.position,
		})),
		{ oldIndex: equations.length, newStart: text.length },
	] as { oldIndex: number; newStart?: number; newEnd?: number }[];
	for (let gap = 0; gap + 1 < bounds.length; gap++) {
		const left = bounds[gap];
		const right = bounds[gap + 1];
		if (!left || !right) continue;
		const missing = equations.slice(left.oldIndex + 1, right.oldIndex);
		const gapStart = left.newEnd ?? 0;
		const gapEnd = right.newStart ?? text.length;
		const spans = mathSpans(text.slice(gapStart, gapEnd)).map((span) => ({
			...span,
			start: span.start + gapStart,
			end: span.end + gapStart,
		}));
		if (missing.length === 0) {
			const moved = spans.find((span) =>
				equations.some(
					(equation) =>
						equation.token === text.slice(span.start, span.end) &&
						!anchors.some((anchor) => anchor.equation === equation),
				),
			);
			if (moved) {
				return ambiguous(
					equations,
					text,
					`${text.slice(moved.start, moved.end)} appears where no equation was — --text can't move an equation`,
				);
			}
			continue; // any span here is literal text
		}
		if (spans.length === 0) {
			deleted.push(...missing);
			continue;
		}
		const paired =
			spans.length === missing.length &&
			spans.every((span, index) => span.display === missing[index]?.display);
		if (!paired) {
			return ambiguous(
				equations,
				text,
				`${describe(missing)} ${missing.length === 1 ? "is" : "are"} missing, but the text in ${missing.length === 1 ? "its" : "their"} place holds ${spans.length} math span${spans.length === 1 ? "" : "s"} (${spans.map((span) => text.slice(span.start, span.end)).join(", ")}) that ${spans.length === 1 ? "doesn't" : "don't"} pair with ${missing.length === 1 ? "it" : "them"} one-to-one by kind`,
			);
		}
		spans.forEach((span, index) => {
			const equation = missing[index];
			if (!equation) return;
			edited.push({ equation, latex: span.latex });
			replacements.push({
				start: span.start,
				end: span.end,
				sentinel: editSentinel(equation.index),
			});
		});
	}

	replacements.sort((left, right) => right.start - left.start);
	let sentineled = text;
	for (const { start, end, sentinel } of replacements) {
		sentineled = sentineled.slice(0, start) + sentinel + sentineled.slice(end);
	}
	return { equations, text: sentineled, edited, deleted };
}

export type OldEquation = {
	index: number;
	node: XmlNode;
	token: string;
	display: boolean;
};

export type EquationTextPlan = {
	/** Every visible equation of the old paragraph, in reading order. */
	equations: OldEquation[];
	/** The new text with kept tokens and edited spans as sentinels. */
	text: string;
	edited: { equation: OldEquation; latex: string }[];
	deleted: OldEquation[];
	/** Set when the rewrite can't be read unambiguously — the caller refuses. */
	ambiguity?: string;
};

/** The paragraph's equations visible in the accepted view, in the order the
 *  `--text` diff walks them (`extractOldTokens`): descending every wrapper but
 *  a deletion/move source, skipping one whose own markers delete it whole. */
export function visibleEquations(paragraph: XmlNode): OldEquation[] {
	const out: OldEquation[] = [];
	function walk(children: XmlNode[]): void {
		for (const child of children) {
			if (isEquation(child)) {
				const token = equationToken(child, "accepted");
				if (token !== null) {
					out.push({
						index: out.length,
						node: child,
						token,
						display: child.tag === "m:oMathPara",
					});
				}
				continue;
			}
			if (
				isRunBearingWrapper(child.tag) &&
				child.tag !== "w:del" &&
				child.tag !== "w:moveFrom"
			) {
				walk(wrapperContent(child));
			}
		}
	}
	walk(paragraph.children);
	return out;
}

/** One diff token for the equation at `index` kept as is. */
export function keepSentinel(index: number): string {
	return `${index}`;
}

/** One diff token for the equation at `index` recompiled from new LaTeX. */
export function editSentinel(index: number): string {
	return `${index}`;
}

/** The equation a sentinel token stands for, or null for an ordinary token. */
export function parseSentinel(
	token: string,
): { kind: "keep" | "edit"; index: number } | null {
	const match = /^([])(\d+)[]$/.exec(token);
	if (!match) return null;
	const [, open, index] = match;
	return { kind: open === "" ? "keep" : "edit", index: Number(index) };
}

/** A sentinel or an ordinary token — what `tokenize` splits so a sentinel is
 *  always ONE token, even glued to punctuation (`($x$)`). */
export const SENTINEL_TOKEN = "\\d+|\\d+";

const SENTINEL_CHARS = /[-]/g;

/** The old equations found verbatim in the new text, in order: each searched
 *  from just past the previous one's match. */
function keptAnchors(
	equations: OldEquation[],
	text: string,
): { equation: OldEquation; position: number }[] {
	const anchors: { equation: OldEquation; position: number }[] = [];
	let cursor = 0;
	for (const equation of equations) {
		const position = text.indexOf(equation.token, cursor);
		if (position === -1) continue;
		// A token found inside a longer `$$…$$` would split it — only take it
		// where it stands as its own span.
		if (!standsAlone(text, position, equation)) continue;
		anchors.push({ equation, position });
		cursor = position + equation.token.length;
	}
	return anchors;
}

function standsAlone(
	text: string,
	position: number,
	equation: OldEquation,
): boolean {
	if (equation.display) return true;
	const before = text[position - 1];
	const after = text[position + equation.token.length];
	return before !== "$" && after !== "$";
}

/** The well-formed math spans in `text`, by Pandoc's `tex_math_dollars` rule
 *  (the stricter reading, so prose dollars aren't mistaken for math):
 *  `$$…$$` with non-empty content, or `$…$` whose opener is followed by a
 *  non-space, whose closer follows a non-space and isn't followed by a digit
 *  (so `$5 and $10` is two prices, not a span). A `\$` is an escaped dollar. */
export function mathSpans(
	text: string,
): { start: number; end: number; latex: string; display: boolean }[] {
	const out: { start: number; end: number; latex: string; display: boolean }[] =
		[];
	let index = 0;
	while (index < text.length) {
		if (text[index] !== "$" || text[index - 1] === "\\") {
			index++;
			continue;
		}
		if (text[index + 1] === "$") {
			const close = text.indexOf("$$", index + 2);
			const latex = close === -1 ? "" : text.slice(index + 2, close);
			if (close !== -1 && latex.trim().length > 0 && !latex.includes("$")) {
				out.push({ start: index, end: close + 2, latex, display: true });
				index = close + 2;
				continue;
			}
			index += 2;
			continue;
		}
		const span = inlineSpanAt(text, index);
		if (span) {
			out.push(span);
			index = span.end;
			continue;
		}
		index++;
	}
	return out;
}

function inlineSpanAt(
	text: string,
	open: number,
): { start: number; end: number; latex: string; display: false } | null {
	const first = text[open + 1];
	if (first === undefined || /\s/.test(first)) return null;
	for (let close = open + 1; close < text.length; close++) {
		if (text[close] !== "$" || text[close - 1] === "\\") continue;
		if (close === open + 1) return null;
		if (/\s/.test(text[close - 1] ?? "")) return null;
		if (/\d/.test(text[close + 1] ?? "")) return null;
		return {
			start: open,
			end: close + 1,
			latex: text.slice(open + 1, close),
			display: false,
		};
	}
	return null;
}

function ambiguous(
	equations: OldEquation[],
	text: string,
	reason: string,
): EquationTextPlan {
	return { equations, text, edited: [], deleted: [], ambiguity: reason };
}

function describe(equations: OldEquation[]): string {
	return equations.map((equation) => equation.token).join(", ");
}
