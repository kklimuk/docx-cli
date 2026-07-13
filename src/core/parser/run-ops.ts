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
		// `<w:delText>` is text content too — runs inside `<w:del>` /
		// `<w:moveFrom>` carry their text in delText. Counting only `<w:t>`
		// here under-reports length for those runs and breaks any caller
		// that walks XML offsets in current view (e.g. find → replace,
		// find → comments add). The AST reader (core/ast/read.ts) already
		// surfaces delText in TextRun.text, so XML-side accounting must
		// agree.
		if (child.tag === "w:t" || child.tag === "w:delText") {
			total += child.collectText().length;
		} else {
			total += inlineMarkerWidth(child);
		}
	}
	return total;
}

/**
 * Slice a `<w:r>` to the offset range `[start, end)`. Four child categories:
 *
 * - `<w:rPr>`: metadata, cloned into every slice.
 * - `<w:t>` / `<w:delText>`: text content, sliced by character offset.
 * - single-character equivalents (`<w:noBreakHyphen>` / `<w:softHyphen>` /
 *   `<w:sym>`) and the whitespace markers `read` renders as one character
 *   (`<w:tab>` / `<w:ptab>` → "\t", `<w:cr>` and a LINE `<w:br>` → "\n"):
 *   own one offset slot, included if their slot lies in `[start, end)`.
 * - other inline children (`<w:drawing>`, `<w:pict>`, `<w:object>`, footnote
 *   refs, and a PAGE/COLUMN `<w:br>` — which `read` omits): zero-width
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
		if (inlineMarkerWidth(child) === 1) {
			if (offset >= start && offset < end) {
				sliced.children.push(child.clone());
			}
			offset += 1;
			continue;
		}
		// Zero-width positional markers (drawings, legacy embeds, note refs, a
		// page/column break): owned by the slice whose range covers this offset.
		if (offset >= start && offset < end) {
			sliced.children.push(child.clone());
		}
	}
	return sliced;
}

/** The offset width of a non-`<w:t>` run child: exactly one character for the
 *  text equivalents (`<w:noBreakHyphen>`/`<w:softHyphen>`/`<w:sym>`) and for the
 *  whitespace markers `read` renders as a character — `<w:tab>`/`<w:ptab>` ("\t"),
 *  `<w:cr>` and a LINE `<w:br>` ("\n"). A PAGE/COLUMN `<w:br>` renders as nothing
 *  in `read`, so it stays zero-width; everything else (drawings, note refs, …) is
 *  a zero-width positional marker. Kept in lockstep with `readRun` in
 *  `core/ast/read.ts` and `paragraphTextForView` in `core/find/index.ts` — the
 *  three must agree or `find` and `replace`/`comments add`/`hyperlinks add`
 *  misalign. */
function inlineMarkerWidth(child: XmlNode): number {
	switch (child.tag) {
		case "w:noBreakHyphen":
		case "w:softHyphen":
		case "w:sym":
		case "w:tab":
		case "w:ptab":
		case "w:cr":
			return 1;
		case "w:br": {
			const type = child.getAttribute("w:type");
			return type === "page" || type === "column" ? 0 : 1;
		}
		default:
			return 0;
	}
}

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

/** Split a paragraph's children into the runs that carry visible content
 *  and everything else (`<w:pPr>`, bookmark markers, comment markers, …).
 *  Run-bearing wrappers (`<w:ins>`, `<w:del>`, `<w:hyperlink>`, etc.) are
 *  flattened — their inner `<w:r>` children surface in the `runs` list and
 *  the wrappers themselves are discarded. This is what every tracked-edit
 *  path needs to do before re-wrapping in fresh `<w:del>`/`<w:ins>`:
 *  without flattening, prior tracked-change wrappers leak alongside the new
 *  ones (the bug from agent feedback: chained edits + hyperlinked paragraphs
 *  duplicating text). */
export function partitionParagraphRuns(paragraph: XmlNode): {
	runs: XmlNode[];
	nonRuns: XmlNode[];
} {
	const runs: XmlNode[] = [];
	const nonRuns: XmlNode[] = [];
	for (const child of paragraph.children) {
		if (child.tag === "w:r") {
			runs.push(child);
			continue;
		}
		if (isRunBearingWrapper(child.tag)) {
			collectInnerRuns(child, runs);
			continue;
		}
		nonRuns.push(child);
	}
	return { runs, nonRuns };
}

function collectInnerRuns(wrapper: XmlNode, out: XmlNode[]): void {
	for (const child of wrapper.children) {
		if (child.tag === "w:r") {
			out.push(child);
			continue;
		}
		if (isRunBearingWrapper(child.tag)) {
			collectInnerRuns(child, out);
		}
	}
}
