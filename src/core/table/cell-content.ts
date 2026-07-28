import type { BlockReference } from "../ast/document/body";
import type { CellReference } from "../locators/resolve";
import { XmlNode } from "../parser";

/** Direct addressable content inside a `<w:tc>`, excluding `<w:tcPr>` and
 * insignificant text nodes. Nested tables remain direct blocks; paragraphs
 * inside them are not flattened into the outer cell. */
export function directCellBlocks(cell: XmlNode): XmlNode[] {
	return cell.children.filter(
		(child) => !child.isText && (child.tag === "w:p" || child.tag === "w:tbl"),
	);
}

/** The cell's sole direct paragraph when it has no nested table or other direct
 * element content. This is the unambiguous shape a bare-cell `edit --at` may
 * alias; richer cells require an explicit `:pK` locator. */
export function soleCellParagraph(cell: XmlNode): XmlNode | null {
	const content = cell.children.filter(
		(child) => !child.isText && child.tag !== "w:tcPr",
	);
	return content.length === 1 && content[0]?.tag === "w:p"
		? (content[0] as XmlNode)
		: null;
}

/** Word emits one mandatory empty paragraph in a blank table cell. Reuse only
 * that conservative shape: `<w:p>` may carry paragraph properties and empty
 * text runs, but no bookmark, field, comment marker, proofing marker, or other
 * meaningful child. */
export function reusableEmptyCellParagraph(cell: XmlNode): XmlNode | null {
	const paragraph = soleCellParagraph(cell);
	if (!paragraph) return null;
	const meaningful = paragraph.children.filter(
		(child) => !child.isText && child.tag !== "w:pPr" && !isEmptyTextRun(child),
	);
	return meaningful.length === 0 ? paragraph : null;
}

function isEmptyTextRun(node: XmlNode): boolean {
	if (node.tag !== "w:r") return false;
	for (const child of node.children) {
		if (child.isText || child.tag === "w:rPr") continue;
		if (child.tag !== "w:t" && child.tag !== "w:delText") return false;
		if (
			child.children.some(
				(text) => !text.isText || (text.text ?? "").length > 0,
			)
		) {
			return false;
		}
	}
	return true;
}

/** Word requires a table cell to contain block content and to end in a
 * paragraph. Apply the deterministic cure after cell-scoped mutation. Returns
 * the synthetic paragraph when one was appended so result-locator reporting can
 * exclude it.
 *
 * Only an EMPTY cell or one ending in a `<w:tbl>` needs the cure. The other
 * block-level children CT_Tc allows in final position (`<w:sdt>`,
 * `<w:customXml>`, `<w:altChunk>`) close with a paragraph of their own, so
 * appending after one would inject a stray blank line into content the caller
 * (today: `docx raw`) only meant to splice into. */
export function ensureCellEndsWithParagraph(parent: XmlNode[]): XmlNode | null {
	const elements = parent.filter(
		(node) => !node.isText && node.tag !== "w:tcPr",
	);
	const last = elements[elements.length - 1];
	if (last && last.tag !== "w:tbl") return null;
	const paragraph = XmlNode.element("w:p");
	parent.push(paragraph);
	return paragraph;
}

/** Apply already-built user blocks at a cell boundary. A standard empty cell
 * reuses its mandatory paragraph when the first result is another paragraph,
 * preserving that node's identity so reread reports the canonical `:p0` handle.
 * The caller repairs the terminal-paragraph invariant after its full mutation
 * pass so batch insertion doesn't append and reposition synthetic paragraphs. */
export function applyCellInsertion(
	cell: XmlNode,
	blocks: XmlNode[],
	mode: "before" | "after",
	reuseParagraph: XmlNode | null,
): XmlNode[] {
	const parent = cell.children;
	const resultNodes: XmlNode[] = [];
	if (reuseParagraph && blocks.length > 0) {
		const first = blocks[0] as XmlNode;
		const anchorIndex = parent.indexOf(reuseParagraph);
		if (anchorIndex === -1) {
			throw new Error("Empty cell paragraph reference is stale");
		}
		if (first.tag === "w:p") {
			reuseParagraph.attributes = {
				...reuseParagraph.attributes,
				...first.attributes,
			};
			reuseParagraph.children = first.children;
			resultNodes.push(reuseParagraph);
			const remaining = blocks.slice(1);
			parent.splice(anchorIndex + 1, 0, ...remaining);
			resultNodes.push(...remaining);
		} else {
			parent.splice(anchorIndex, 1, ...blocks);
			resultNodes.push(...blocks);
		}
	} else {
		const content = directCellBlocks(cell);
		const boundary =
			mode === "before" ? content[0] : content[content.length - 1];
		const boundaryIndex = boundary ? parent.indexOf(boundary) : parent.length;
		const insertIndex =
			mode === "after" && boundary ? boundaryIndex + 1 : boundaryIndex;
		parent.splice(insertIndex, 0, ...blocks);
		resultNodes.push(...blocks);
	}
	return resultNodes;
}

/** The paragraph a cell insert inherits its formatting from: the reused
 * mandatory empty paragraph when there is one, else the cell's first or last
 * direct paragraph depending on which boundary we're inserting at. Returns null
 * for a malformed cell with no direct paragraph at all. */
export function cellInsertionAnchor(
	cell: CellReference,
	mode: "before" | "after",
	reusable: XmlNode | null,
): BlockReference | null {
	if (reusable) return { node: reusable, parent: cell.parent };
	return mode === "before"
		? (cell.paragraphs[0] ?? null)
		: (cell.paragraphs[cell.paragraphs.length - 1] ?? null);
}
