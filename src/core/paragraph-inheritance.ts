import { insertPprChildInOrder, overlayParagraphAttributes } from "./blocks";
import { partitionParagraphRuns, type XmlNode } from "./parser";

/** Whether a replacement paragraph brings block structure the paragraph it
 * replaces did NOT have — a markdown `# heading` landing on a body line, or a
 * list item landing on a plain paragraph. Such a paragraph owns its look through
 * its style and must not inherit the old pPr/rPr (an 8pt Arial form cell
 * replaced with `## Heading` would otherwise render the heading at 8pt Arial).
 * The comparison is against the OLD paragraph on purpose: a `- item` replacing a
 * `ListParagraph` bullet, or `## X` replacing another Heading2, is the SAME
 * structure and inherits like plain text — judged only on the new paragraph, a
 * bullet retyped through `--markdown` lost its Calibri, tab stops, spacing and
 * indent (batch-3 résumé), which is exactly what steering agents to markdown
 * must not cost. */
export function bringsNewBlockStructure(
	newParagraph: XmlNode,
	oldParagraph: XmlNode,
): boolean {
	const newPpr = newParagraph.findChild("w:pPr");
	const oldPpr = oldParagraph.findChild("w:pPr");
	const newStyle = newPpr?.findChild("w:pStyle")?.getAttribute("w:val");
	const oldStyle = oldPpr?.findChild("w:pStyle")?.getAttribute("w:val");
	const newList = Boolean(newPpr?.findChild("w:numPr"));
	const oldList = Boolean(oldPpr?.findChild("w:numPr"));
	// A list item replacing an UNSTYLED list item is the same structure too: the
	// markdown walker stamps `ListParagraph` on every `- item`, while Google Docs
	// and LibreOffice bullets are a bare `<w:numPr>` — comparing styles alone
	// called that "new" and dropped the bullet's font, tabs and indent.
	if (newList && oldList && !oldStyle) return false;
	if (newStyle && newStyle !== oldStyle) return true;
	return newList && !oldList;
}

/** Give every plain new paragraph the old paragraph's properties, with explicit
 * new values winning. Shared by whole-paragraph edit and empty-cell insert reuse
 * so direct alignment/spacing/indent and paragraph-mark formatting survive.
 * Reusing the SAME cell paragraph preserves an existing pPrChange; replacement
 * paragraphs drop it because they get their own tracked snapshot when needed. */
export function inheritParagraphFormattingIfPlain(
	oldParagraph: XmlNode,
	newParagraphs: XmlNode[],
	explicitStyle: string | undefined,
	options: { preservePprChange?: boolean } = {},
): void {
	const oldPpr = oldParagraph.findChild("w:pPr");
	if (!oldPpr) return;
	for (const newParagraph of newParagraphs) {
		if (newParagraph.tag !== "w:p") continue;
		if (bringsNewBlockStructure(newParagraph, oldParagraph)) continue;
		mergeInheritedPpr(
			oldPpr,
			newParagraph,
			explicitStyle,
			Boolean(options.preservePprChange),
		);
	}
}

/** pPr children whose attributes are independent settings (merge them) rather
 * than one value (replace it). */
export const ATTRIBUTE_BAG_TAGS: ReadonlySet<string> = new Set([
	"w:spacing",
	"w:ind",
]);

/** An inherited attribute bag with the fresh element's attributes laid over
 * it (`overlayParagraphAttributes`). */
export function overlayAttributes(inherited: XmlNode, fresh: XmlNode): XmlNode {
	const merged = inherited.clone();
	overlayParagraphAttributes(merged, fresh.attributes);
	return merged;
}

/** Clone the old `<w:pPr>` onto one replacement paragraph, letting anything the
 * new paragraph already set win while preserving canonical CT_PPr order. */
function mergeInheritedPpr(
	oldPpr: XmlNode,
	newParagraph: XmlNode,
	explicitStyle: string | undefined,
	preservePprChange: boolean,
): void {
	const merged = oldPpr.clone();
	merged.children = merged.children.filter(
		(child) =>
			child.tag !== "w:sectPr" &&
			(preservePprChange || child.tag !== "w:pPrChange"),
	);
	if (explicitStyle) {
		merged.children = merged.children.filter(
			(child) => child.tag !== "w:pStyle",
		);
	}
	const newPpr = newParagraph.findChild("w:pPr");
	if (newPpr) {
		for (const child of newPpr.children) {
			const index = merged.children.findIndex((own) => own.tag === child.tag);
			if (index < 0) {
				insertPprChildInOrder(merged, child);
				continue;
			}
			const existing = merged.children[index] as XmlNode;
			// `<w:spacing>`/`<w:ind>` are attribute BAGS — a new `space-after` must
			// not throw away the inherited `before`/`line`, nor a new `indent-left`
			// the inherited `right`. Overlay attribute-by-attribute (the same rule
			// `applyParagraphPropsToPPr` follows in place), honoring the one
			// exclusive pair: firstLine and hanging share a slot. Every other pPr
			// child is a single value and the new one simply wins.
			merged.children[index] = ATTRIBUTE_BAG_TAGS.has(child.tag)
				? overlayAttributes(existing, child)
				: child;
		}
		newParagraph.children[newParagraph.children.indexOf(newPpr)] = merged;
		return;
	}
	newParagraph.children.unshift(merged);
}

/** A paragraph's `<w:pStyle>` id, if any. */
export function paragraphStyleId(paragraph: XmlNode): string | undefined {
	return paragraph
		.findChild("w:pPr")
		?.findChild("w:pStyle")
		?.getAttribute("w:val");
}

/** Paragraph styles that set their own size and look on purpose (headings,
 * title, TOC levels) — new content neither lends them a body line's look nor
 * shrinks them to body size. */
export function isHeadingLikeStyle(styleId: string | undefined): boolean {
	return styleId !== undefined && HEADING_LIKE.test(styleId);
}

const HEADING_LIKE = /^(heading[1-9]|title|subtitle|toc)/i;

/** The runs of a paragraph that carry text (`<w:t>`), wrappers resolved. */
export function textRuns(paragraph: XmlNode): XmlNode[] {
	return partitionParagraphRuns(paragraph).runs.filter((run) =>
		run.findChild("w:t"),
	);
}

/** Every `<w:p>` in freshly-built content, descending into tables a markdown
 * source may have produced. */
export function paragraphsIn(nodes: XmlNode[]): XmlNode[] {
	const out: XmlNode[] = [];
	for (const node of nodes) {
		if (node.tag === "w:p") out.push(node);
		else if (node.children.length > 0) out.push(...paragraphsIn(node.children));
	}
	return out;
}
