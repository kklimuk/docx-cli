import type { BlockReference } from "../ast/document/body";
import { CellTargetError } from "../locators/cell-target-error";
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
			// A `CellTargetError`, not a bare Error: every cell-insertion caller
			// already translates this class into a `{code, error}` payload, so a
			// detached ref reports as a normal failure instead of a stack trace.
			throw new CellTargetError(
				"BLOCK_NOT_FOUND",
				"Empty cell paragraph reference is stale",
			);
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

/** Where successive inserts land inside ONE cell.
 *
 * {@link applyCellInsertion} re-derives the boundary from the cell's CURRENT
 * blocks. That is already right for every append (the cell's last block IS the
 * one the previous entry appended) and for the FIRST prepend. It is wrong in
 * exactly one case: a SECOND prepend, which would re-derive to the block the
 * first one just placed and land ahead of it, reversing JSONL order.
 *
 * So the only state a batch actually needs is a start cursor, and this class
 * owns that rule rather than each batch surface re-deriving it. The cursor is a
 * live `XmlNode` ref into the cell's child list, so an instance is valid only
 * while that list is the one being spliced (the batch splice phase). A detached
 * ref raises a stale-reference `CellTargetError` rather than silently
 * repositioning. */
export class CellInsertionCursor {
	/** The last block a `before` entry placed; null until one has. */
	private beforeCursor: XmlNode | null = null;

	private constructor(private readonly cell: CellReference) {}

	/** Open a cursor over a cell's direct blocks. Refuses a cell with none —
	 * Word always leaves one mandatory paragraph, so that shape is malformed
	 * rather than one we should invent a boundary for. */
	static open(cell: CellReference): CellInsertionCursor {
		if (cell.blocks.length === 0) {
			throw new CellTargetError(
				"TABLE_STRUCTURE",
				`Cell ${cell.id} has no direct block content`,
			);
		}
		return new CellInsertionCursor(cell);
	}

	/** Place already-built blocks at this cell's start or end boundary. Returns
	 * the nodes now in the tree — for a `reuseParagraph` insert that includes the
	 * REUSED node, which keeps its identity (and so its canonical `:p0` handle). */
	insert(
		blocks: XmlNode[],
		mode: "before" | "after",
		reuseParagraph: XmlNode | null,
	): XmlNode[] {
		if (blocks.length === 0) {
			throw new CellTargetError(
				"TABLE_STRUCTURE",
				`Insertion into ${this.cell.id} produced no addressable blocks`,
			);
		}
		// Everything but a repeat prepend can re-derive its own boundary.
		if (mode === "after" || reuseParagraph || !this.beforeCursor) {
			const applied = applyCellInsertion(
				this.cell.node,
				blocks,
				mode,
				reuseParagraph,
			);
			if (mode === "before") {
				this.beforeCursor = applied[applied.length - 1] ?? null;
			}
			return applied;
		}
		// Repeat prepend: stack forward from the last one so entry order survives.
		const index = this.cell.parent.indexOf(this.beforeCursor);
		if (index === -1) {
			throw new CellTargetError(
				"BLOCK_NOT_FOUND",
				`Cell boundary reference is stale for ${this.cell.id}`,
			);
		}
		this.cell.parent.splice(index + 1, 0, ...blocks);
		this.beforeCursor = blocks[blocks.length - 1] ?? null;
		return blocks;
	}

	/** Apply Word's terminal-paragraph rule once the cell's whole batch is in.
	 * Deferred to the end so a synthetic paragraph isn't appended and then
	 * repositioned by a later entry. */
	ensureTerminalParagraph(): void {
		ensureCellEndsWithParagraph(this.cell.parent);
	}
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
