import { XmlNode } from "../parser/xml-node";
import { MATH_REVISION_TAGS } from "./handlers";
import { ommlToLatex } from "./read";

/** The token an equation occupies in the find/replace/span text: exactly what
 *  `read` prints for it — `$latex$` inline, `$$latex$$` for a display block —
 *  or null when it occupies nothing (no math text at all, which the reader
 *  skips, or a whole-equation revision `view` hides: Word's per-math-run
 *  `<w:del>` markers in the accepted view, `<w:ins>` ones in the baseline).
 *  An enclosing `<w:del>`/`<w:ins>` WRAPPER is the walker's business — every
 *  offset walker already skips a wrapper its view hides — so only the markers
 *  INSIDE the equation are decided here.
 *
 *  This is the XML-side twin of the AST's `EquationRun` (`read.ts` builds the
 *  run from `equationLatex`/`mathText` below), so both sides of the offset
 *  bridge give one equation the same width. */
export function equationToken(
	node: XmlNode,
	view: TokenView = "current",
): string | null {
	const text = mathText(node);
	if (text.length === 0) return null;
	const { whole } = mathRevisionCoverage(node);
	if (whole === "w:del" && view === "accepted") return null;
	if (whole === "w:ins" && view === "baseline") return null;
	return formatEquationToken(
		equationLatex(node),
		text,
		node.tag === "m:oMathPara",
	);
}

export type TokenView = "current" | "accepted" | "baseline";

/** `$latex$` / `$$latex$$` — the one spelling `read`'s markdown, `find`'s
 *  search text and the XML walkers share. `text` (the flat math text) stands
 *  in when the LaTeX walker produced nothing, as it does in `read`. */
export function formatEquationToken(
	latex: string,
	text: string,
	display: boolean,
): string {
	const body = latex.length > 0 ? latex : text;
	return display ? `$$${body}$$` : `$${body}$`;
}

/** The LaTeX `read` shows for an equation. A PARTLY revised equation (Word's
 *  per-math-run markers covering only some of it) reads as its accepted view —
 *  what `accept --all` leaves — not the deleted symbols beside the kept ones. */
export function equationLatex(node: XmlNode): string {
	const { partial } = mathRevisionCoverage(node);
	return ommlToLatex(partial ? withoutMathDeletions(node) : node);
}

/** Concatenate all `<m:t>` descendants in document order. The structural OMML
 *  (subs, sups, fractions) collapses to flat characters — degraded but
 *  readable; an equation with none is no equation at all. */
export function mathText(node: XmlNode): string {
	let out = "";
	for (const child of node.children) {
		if (child.tag === "m:t") {
			out += child.collectText();
			continue;
		}
		out += mathText(child);
	}
	return out;
}

/** How Word's in-equation revision markers cover it: `whole` names the marker
 *  tag when ONE revision (one kind, one author) covers every symbol, the way
 *  an enclosing wrapper would; `partial` is set when markers exist but don't.
 *  Grouped by tag + author exactly as `registerMathRevisions` groups them into
 *  `tcN`s, so the reader and the walkers agree on which is which. */
export function mathRevisionCoverage(node: XmlNode): {
	whole?: "w:ins" | "w:del";
	partial: boolean;
} {
	const groups = new Map<string, "w:ins" | "w:del">();
	collectMarkerGroups(node, groups);
	if (groups.size === 0) return { partial: false };
	const [only] = groups.values();
	if (groups.size === 1 && only && !hasUnmarkedMathText(node)) {
		return { whole: only, partial: false };
	}
	return { partial: true };
}

function collectMarkerGroups(
	node: XmlNode,
	groups: Map<string, "w:ins" | "w:del">,
): void {
	for (const child of node.children) {
		if (!MATH_REVISION_TAGS.has(child.tag)) {
			collectMarkerGroups(child, groups);
			continue;
		}
		const tag = child.tag === "w:ins" ? "w:ins" : "w:del";
		groups.set(`${child.tag} ${child.getAttribute("w:author") ?? ""}`, tag);
	}
}

/** Whether any of the equation's text sits outside a revision marker. */
export function hasUnmarkedMathText(node: XmlNode): boolean {
	return node.children.some((child) => {
		if (MATH_REVISION_TAGS.has(child.tag)) return false;
		if (child.tag === "m:t") return child.collectText().length > 0;
		return hasUnmarkedMathText(child);
	});
}

/** A copy of `node` as it reads once its tracked deletions are accepted — a
 *  copy, so the live tree keeps every marker for `accept`/`reject`. */
function withoutMathDeletions(node: XmlNode): XmlNode {
	const copy = new XmlNode(
		node.tag,
		node.attributes,
		node.children
			.filter((child) => child.tag !== "w:del")
			.map(withoutMathDeletions),
	);
	if (node.text !== undefined) copy.text = node.text;
	return copy;
}
