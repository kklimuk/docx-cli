import { XmlNode } from "./xml-node";

/** Wrappers whose contents store text as `<w:delText>` rather than `<w:t>`.
 *  Matters when a span replace cuts content out of these wrappers — we don't
 *  need to re-wrap the cut in a new `<w:del>` (it's already deleted), and on
 *  reject we need to rename `<w:delText>` → `<w:t>` before unwrapping. */
export function isSubtractiveTrackedChangeWrapper(tag: string): boolean {
	return tag === "w:del" || tag === "w:moveFrom";
}

export function runTextLength(run: XmlNode): number {
	let total = 0;
	for (const child of run.children) {
		if (child.tag === "w:t") total += child.collectText().length;
		else if (SINGLE_CHAR_TAGS.has(child.tag)) total += 1;
	}
	return total;
}

/**
 * Slice a `<w:r>` to the offset range `[start, end)`. Four child categories:
 *
 * - `<w:rPr>`: metadata, cloned into every slice.
 * - `<w:t>` / `<w:delText>`: text content, sliced by character offset.
 * - `<w:noBreakHyphen>` / `<w:softHyphen>` / `<w:sym>`: single-character
 *   text equivalents — own one offset slot, included if their slot lies in
 *   `[start, end)`.
 * - other inline children (`<w:tab>`, `<w:br>`, `<w:cr>`, `<w:ptab>`,
 *   `<w:drawing>`, `<w:pict>`, `<w:object>`, footnote refs): zero-width
 *   positional markers — owned by the slice satisfying
 *   `start <= offset < end`. This partitions cleanly across pre/cut/post.
 */
export function sliceRun(run: XmlNode, start: number, end: number): XmlNode {
	const sliced = new XmlNode("w:r", { ...run.attributes });
	let offset = 0;
	for (const child of run.children) {
		if (child.tag === "w:rPr") {
			sliced.children.push(child.clone());
			continue;
		}
		if (child.tag === "w:t" || child.tag === "w:delText") {
			const text = child.collectText();
			const localStart = Math.max(0, start - offset);
			const localEnd = Math.min(text.length, end - offset);
			if (localStart < localEnd) {
				const slicedText = new XmlNode(child.tag, { "xml:space": "preserve" });
				slicedText.children.push(
					XmlNode.textNode(text.slice(localStart, localEnd)),
				);
				sliced.children.push(slicedText);
			}
			offset += text.length;
			continue;
		}
		if (SINGLE_CHAR_TAGS.has(child.tag)) {
			if (offset >= start && offset < end) {
				sliced.children.push(child.clone());
			}
			offset += 1;
			continue;
		}
		if (offset >= start && offset < end) {
			sliced.children.push(child.clone());
		}
	}
	return sliced;
}

/** Tags that contribute exactly one character to the AST text and to
 *  paragraph-level offset accounting. Kept in sync with `readRun`'s handling
 *  of these elements in `core/ast/read.ts`. */
const SINGLE_CHAR_TAGS = new Set(["w:noBreakHyphen", "w:softHyphen", "w:sym"]);

/** Paragraph-level wrappers whose inner runs contribute to the paragraph's
 *  text content. Anything that `walkRunContainer` in `core/ast/read.ts`
 *  recurses into (besides `<w:r>` itself) must appear here so the AST text
 *  and the XML-side offset arithmetic agree. The two have to drift together
 *  or `find` and `replace`/`comments add`/`hyperlinks add` will misalign.
 *
 *  - `w:ins` / `w:del` / `w:moveFrom` / `w:moveTo`: tracked-change wrappers.
 *  - `w:hyperlink`: hyperlink span (own a relationship, runs are visible text).
 *  - `w:fldSimple`: self-contained field; runs render the cached field result.
 *  - `w:smartTag`: semantic annotation around runs (person names, dates).
 */
export const RUN_BEARING_WRAPPER_TAGS: ReadonlySet<string> = new Set([
	"w:ins",
	"w:del",
	"w:moveFrom",
	"w:moveTo",
	"w:hyperlink",
	"w:fldSimple",
	"w:smartTag",
]);

export function isRunBearingWrapper(tag: string): boolean {
	return RUN_BEARING_WRAPPER_TAGS.has(tag);
}

/** Sum the text lengths of all `<w:r>` reachable from `children`, descending
 *  transparently through every run-bearing wrapper. Matches what the AST's
 *  `paragraph.runs.map(r => r.text).join("").length` would compute, so this
 *  is the canonical XML-side counterpart for `find`'s offset arithmetic. */
export function sumRunBearingTextLength(children: XmlNode[]): number {
	let total = 0;
	for (const child of children) {
		if (child.tag === "w:r") {
			total += runTextLength(child);
			continue;
		}
		if (isRunBearingWrapper(child.tag)) {
			total += sumRunBearingTextLength(child.children);
		}
	}
	return total;
}
