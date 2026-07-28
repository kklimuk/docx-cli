import { insertPprChildInOrder } from "./blocks";
import type { XmlNode } from "./parser";

/** A paragraph that brings its own style/list structure owns its block look and
 * must not inherit a replaced/placeholder paragraph's pPr. */
export function paragraphOwnsBlockStructure(paragraph: XmlNode): boolean {
	const ownPpr = paragraph.findChild("w:pPr");
	return Boolean(ownPpr?.findChild("w:pStyle") || ownPpr?.findChild("w:numPr"));
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
		if (paragraphOwnsBlockStructure(newParagraph)) continue;
		mergeInheritedPpr(
			oldPpr,
			newParagraph,
			explicitStyle,
			Boolean(options.preservePprChange),
		);
	}
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
			if (index >= 0) merged.children[index] = child;
			else insertPprChildInOrder(merged, child);
		}
		newParagraph.children[newParagraph.children.indexOf(newPpr)] = merged;
		return;
	}
	newParagraph.children.unshift(merged);
}
