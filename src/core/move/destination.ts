import type { Document } from "../ast/document";
import {
	type CellReference,
	parseCellAt,
	resolveCellReference,
} from "../locators/resolve";
import type { XmlNode } from "../parser";
import { reusableEmptyCellParagraph } from "../table";
import { MoveError } from "./move-error";
import { storyOf } from "./source";

/** Where the moved blocks land — the same two shapes `insert --before/--after`
 *  takes: beside a block (`pN`, `tN`, a cell or text-box paragraph), or at a
 *  bare cell's start (`before`) / end (`after`). */
export type MoveDestination =
	| {
			kind: "block";
			locator: string;
			side: "before" | "after";
			anchor: XmlNode;
			parent: XmlNode[];
			story: string;
	  }
	| {
			kind: "cell";
			locator: string;
			side: "before" | "after";
			cell: CellReference;
			/** Word's mandatory empty paragraph, when that's all the cell holds —
			 *  the first move in fills it instead of stacking beside it. */
			reusableParagraph: XmlNode | null;
			story: string;
	  };

export function resolveMoveDestination(
	document: Document,
	locator: string,
	side: "before" | "after",
): MoveDestination {
	const story = storyOf(locator);
	if (parseCellAt(locator)) {
		const cell = resolveCellReference(document, locator);
		return {
			kind: "cell",
			locator,
			side,
			cell,
			reusableParagraph: reusableEmptyCellParagraph(cell.node),
			story,
		};
	}
	const reference = document.body.resolveBlock(locator);
	if (reference.node.tag !== "w:p" && reference.node.tag !== "w:tbl") {
		throw new MoveError(
			"INVALID_LOCATOR",
			`${locator} can't anchor a move — use the paragraph or table beside it`,
			"Anchor on a paragraph or table (pN, tN, tN:rRcC:pK, tbxN:pK) or a bare cell (tN:rRcC). For the very end of the document, use --after on its last block.",
		);
	}
	return {
		kind: "block",
		locator,
		side,
		anchor: reference.node,
		parent: reference.parent,
		story,
	};
}
