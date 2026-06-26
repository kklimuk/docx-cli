import type { Document } from "../ast/document";
import type { BlockReference } from "../ast/document/body";
import {
	FORMAT_TO_NUMFMT,
	type ListFormat,
	type NumberingView,
} from "../ast/document/numbering";
import type { XmlNode } from "../parser";

/** Cross-cutting lens over a document's list numbering. `--start`/`--format`
 * are pure `numbering.xml` mutations keyed by a paragraph's `numId`/`ilvl`
 * (they affect the whole list); `--restart`/`--continue` ALSO re-point the
 * `<w:numId>` on the body paragraphs, so the lens spans `numbering` + the body.
 *
 * In our model each `allocate()` mints a unique `numId`, so "the list" a
 * paragraph belongs to is every sibling carrying that same `numId`. The
 * restart/continue walks are scoped to the paragraph's PARENT sibling array
 * (the body, or one table cell) — never `Body.iterateBlocks`, which flattens
 * across cell boundaries and would make "preceding"/"following" ill-defined. */
export class Lists {
	constructor(private document: Document) {}

	/** Set the start value of the addressed paragraph's list. Affects the whole
	 * list (every paragraph sharing its numId). */
	setStart(blockRef: BlockReference, start: number): void {
		const list = this.requireList(blockRef.node);
		this.numbering().setStart(list.numId, list.level, start);
	}

	/** Override the numbering format (decimal / lower-alpha / upper-roman / …) of
	 * the addressed paragraph's list, at that paragraph's level. */
	setFormat(blockRef: BlockReference, format: ListFormat): void {
		const list = this.requireList(blockRef.node);
		this.numbering().setFormat(
			list.numId,
			list.level,
			FORMAT_TO_NUMFMT[format],
		);
	}

	/** Begin a fresh list at the addressed paragraph: mint a new numId (copying
	 * the current format, starting at `start`) and re-point this paragraph plus
	 * every LATER sibling currently sharing the old numId. Earlier same-numId
	 * siblings keep the old list. */
	restart(blockRef: BlockReference, start: number): void {
		const list = this.requireList(blockRef.node);
		const newNumId = String(
			this.numbering().cloneListDefinition(list.numId, start),
		);
		const startIdx = blockRef.parent.indexOf(blockRef.node);
		this.repointFrom(blockRef.parent, startIdx, list.numId, newNumId);
	}

	/** Continue the immediately-preceding list's numbering: re-point this
	 * paragraph's list run to the nearest earlier ordered list's numId, so the
	 * shared numId makes Word (and the markdown ordinal) count across both. */
	continue(blockRef: BlockReference): void {
		const list = this.requireList(blockRef.node);
		const siblings = blockRef.parent;
		const startIdx = siblings.indexOf(blockRef.node);
		let precedingNumId: string | undefined;
		for (let i = startIdx - 1; i >= 0; i--) {
			const sibling = siblings[i];
			if (!sibling) continue;
			const numbered = readNumPr(sibling);
			if (!numbered) continue;
			if (numbered.numId !== list.numId && this.isOrdered(numbered)) {
				precedingNumId = numbered.numId;
				break;
			}
		}
		if (!precedingNumId) {
			throw new ListOperationError("no preceding list to continue from");
		}
		this.repointFrom(siblings, startIdx, list.numId, precedingNumId);
	}

	/** Re-point the run at `startIdx` and every later sibling carrying `oldNumId`
	 * to `newNumId` (rewriting the existing `<w:numId w:val>` in place). */
	private repointFrom(
		siblings: XmlNode[],
		startIdx: number,
		oldNumId: string,
		newNumId: string,
	): void {
		for (let i = startIdx; i < siblings.length; i++) {
			const sibling = siblings[i];
			if (!sibling) continue;
			if (readNumPr(sibling)?.numId === oldNumId) setNumId(sibling, newNumId);
		}
	}

	private isOrdered(numbered: { numId: string; level: number }): boolean {
		const format = this.numbering().getFormat(numbered.numId, numbered.level);
		return format !== undefined && format !== "bullet" && format !== "none";
	}

	private requireList(paragraph: XmlNode): { numId: string; level: number } {
		const numbered = readNumPr(paragraph);
		if (!numbered) {
			throw new ListOperationError("paragraph is not a list item");
		}
		return numbered;
	}

	private numbering(): NumberingView {
		if (!this.document.numbering) {
			throw new ListOperationError("document has no numbering definitions");
		}
		return this.document.numbering;
	}
}

/** Read a paragraph's `<w:pPr><w:numPr>` numId + level, or undefined if it
 * carries no numbering. Mirrors the reader at `core/ast/read.ts`. */
function readNumPr(
	paragraph: XmlNode,
): { numId: string; level: number } | undefined {
	const numPr = paragraph.findChild("w:pPr")?.findChild("w:numPr");
	if (!numPr) return undefined;
	const numId = numPr.findChild("w:numId")?.getAttribute("w:val");
	if (numId === undefined) return undefined;
	const level = Number(numPr.findChild("w:ilvl")?.getAttribute("w:val") ?? "0");
	return { numId, level };
}

function setNumId(paragraph: XmlNode, numId: string): void {
	paragraph
		.findChild("w:pPr")
		?.findChild("w:numPr")
		?.findChild("w:numId")
		?.setAttribute("w:val", numId);
}

export class ListOperationError extends Error {
	constructor(message: string) {
		super(message);
		this.name = "ListOperationError";
	}
}
