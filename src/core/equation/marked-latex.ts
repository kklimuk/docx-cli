import { XmlNode } from "../parser";
import { MATH_REVISION_TAGS } from "./handlers";
import { ommlToLatex } from "./read";

/** The LaTeX of just the part of an equation that `markers` mark — Word's
 *  per-math-run `<w:del>`/`<w:ins>` for one revision. A partial deletion of
 *  `E=mc^2` that took `c^2` reads `c^{2}`, not the joined `c2`: the structure
 *  (here the `<m:sSup>`) survives around the marked symbols, so the text can be
 *  copied back into `equations edit`. Built from a pruned COPY (the live tree
 *  keeps every marker for accept/reject): unmarked symbols drop, a structure
 *  left with no marked symbol drops, and a structure's argument slots and
 *  property elements stay so its handler still sees a well-formed shape. */
export function markedLatex(
	equation: XmlNode,
	markers: ReadonlySet<XmlNode>,
): string {
	const pruned = pruneToMarked(equation, markers, true);
	return pruned ? ommlToLatex(pruned) : "";
}

function pruneToMarked(
	node: XmlNode,
	markers: ReadonlySet<XmlNode>,
	keepEmpty: boolean,
): XmlNode | undefined {
	if (markers.has(node)) return node;
	// Another revision's marker, or a symbol no marker of this one covers.
	if (MATH_REVISION_TAGS.has(node.tag) || node.tag === "m:t" || node.isText)
		return undefined;
	if (node.tag.endsWith("Pr")) return node;
	const children: XmlNode[] = [];
	for (const child of node.children) {
		const kept = pruneToMarked(child, markers, ARGUMENT_SLOTS.has(child.tag));
		if (kept) children.push(kept);
	}
	const copy = new XmlNode(node.tag, node.attributes, children);
	if (!keepEmpty && copy.collectText().length === 0) return undefined;
	return copy;
}

/** A structure's argument slots: kept even when empty, so `<m:sSup>` with only
 *  its exponent marked still reads `{}^{2}` rather than losing its shape. */
const ARGUMENT_SLOTS = new Set([
	"m:e",
	"m:sup",
	"m:sub",
	"m:num",
	"m:den",
	"m:deg",
	"m:fName",
	"m:lim",
]);
