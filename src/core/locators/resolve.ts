import type { Document } from "../ast/document";
import {
	type BlockRangeReference,
	type BlockReference,
	LocatorResolveError,
} from "../ast/document/body";
import type { XmlNode } from "../parser";
import { directCellBlocks, soleCellParagraph } from "../table/cell-content";
import { buildGrid, cellAt, type Grid, resolveTableNode } from "../table/grid";
import { CellTargetError } from "./cell-target-error";
import { type Locator, LocatorParseError, parseLocator } from "./parse";

export { type BlockRangeReference, LocatorResolveError };

export type BlockTarget = {
	blockId: string;
	span?: { start: number; end: number };
};

/**
 * Flatten a locator that targets a single block (possibly nested in a table cell)
 * into the registered `blockId` plus an optional `span`. Returns `null` for
 * locators that don't address a single block (e.g. cross-block ranges, comments,
 * images, whole-cell locators without an inner block).
 */
export function locatorToBlockTarget(locator: Locator): BlockTarget | null {
	if (locator.kind === "block") return { blockId: locator.blockId };
	if (locator.kind === "blockSpan") {
		return {
			blockId: locator.blockId,
			span: { start: locator.start, end: locator.end },
		};
	}
	if (locator.kind === "cell" && locator.inner) {
		const inner = locatorToBlockTarget(locator.inner);
		if (!inner) return null;
		return {
			blockId: `${locator.tableId}:r${locator.row}c${locator.col}:${inner.blockId}`,
			span: inner.span,
		};
	}
	return null;
}

export type CellReference = {
	id: string;
	tableId: string;
	row: number;
	col: number;
	node: XmlNode;
	parent: XmlNode[];
	blocks: XmlNode[];
	paragraphs: BlockReference[];
};

/** Resolve a bare `tN:rRcC` (including nested chains) to its physical cell.
 * Bare-cell content mutation deliberately rejects merged/grid-shifted rows:
 * GFM cannot faithfully show their logical grid and writing to a vMerge
 * continuation may succeed invisibly. Explicit `:pK` locators retain their
 * existing behavior for callers that intentionally address such structures.
 *
 * `gridCache` is for callers that resolve MANY cells against one unmutated tree
 * (`insert`/`edit --batch`): `buildGrid` walks every row and cell of the table,
 * so a 40-entry form-fill of one table would otherwise rebuild the same grid 40
 * times. A single-shot resolve passes nothing and builds once, as before. */
export function resolveCellReference(
	document: Document,
	at: string,
	gridCache?: Map<XmlNode, Grid>,
): CellReference {
	const target = parseCellAt(at);
	if (!target) {
		throw new CellTargetError(
			"INVALID_LOCATOR",
			`"${at}" is not a bare table-cell locator`,
			"Use a cell locator such as t0:r1c2. Paragraph locators still use --before/--after or edit --at t0:r1c2:p0.",
		);
	}
	const table = resolveTableNode(document, target.tableId);
	if (!table) {
		throw new CellTargetError(
			"BLOCK_NOT_FOUND",
			`Table not found: ${target.tableId}`,
		);
	}
	const grid = cachedGrid(table, gridCache);
	const row = grid.rows[target.row];
	if (!row) {
		throw new CellTargetError(
			"BLOCK_NOT_FOUND",
			`Row r${target.row} does not exist in ${target.tableId}`,
		);
	}
	const trPr = row.node.findChild("w:trPr");
	if (
		trPr?.findChild("w:gridBefore") ||
		trPr?.findChild("w:gridAfter") ||
		row.cells.some((cell) => cell.colSpan !== 1)
	) {
		throw mergedCellTargetError(at);
	}
	const cell = cellAt(row, target.col);
	if (!cell) {
		throw new CellTargetError(
			"BLOCK_NOT_FOUND",
			`Cell r${target.row}c${target.col} does not exist in ${target.tableId}`,
		);
	}
	if (cell.vMerge || cell.colSpan !== 1) throw mergedCellTargetError(at);
	const unsupported = cell.node.children.find(
		(child) =>
			!child.isText &&
			child.tag !== "w:tcPr" &&
			child.tag !== "w:p" &&
			child.tag !== "w:tbl",
	);
	if (unsupported) {
		throw new CellTargetError(
			"TABLE_STRUCTURE",
			`Cell ${at} contains unsupported direct ${unsupported.tag} content`,
			"Target an explicit paragraph locator inside the cell instead.",
		);
	}
	const id = `${target.tableId}:r${target.row}c${cell.colStart}`;
	const paragraphs = cell.node.findChildren("w:p").map((paragraph) => ({
		node: paragraph,
		parent: cell.node.children,
	}));
	return {
		id,
		tableId: target.tableId,
		row: target.row,
		col: cell.colStart,
		node: cell.node,
		parent: cell.node.children,
		blocks: directCellBlocks(cell.node),
		paragraphs,
	};
}

function cachedGrid(
	table: XmlNode,
	cache: Map<XmlNode, Grid> | undefined,
): Grid {
	if (!cache) return buildGrid(table);
	const cached = cache.get(table);
	if (cached) return cached;
	const grid = buildGrid(table);
	cache.set(table, grid);
	return grid;
}

/** Resolve the unambiguous cell shape accepted by `edit --at CELL`: exactly one
 * direct paragraph and no nested/direct sibling blocks. */
export function resolveCellParagraphReference(
	document: Document,
	at: string,
	gridCache?: Map<XmlNode, Grid>,
): { cell: CellReference; paragraph: BlockReference } {
	const cell = resolveCellReference(document, at, gridCache);
	const paragraph = soleCellParagraph(cell.node);
	if (!paragraph) {
		throw new CellTargetError(
			"TABLE_STRUCTURE",
			`Cell ${at} has multiple blocks; a bare cell edit would be ambiguous`,
			`Re-read the document and target an explicit paragraph such as ${cell.id}:p0.`,
		);
	}
	return {
		cell,
		paragraph: { node: paragraph, parent: cell.parent },
	};
}

function mergedCellTargetError(at: string): CellTargetError {
	return new CellTargetError(
		"TABLE_STRUCTURE",
		`Bare-cell content mutation is not supported for merged or grid-shifted cell ${at}`,
		"Run `docx read FILE` and use the tN:rRcC:pK paragraph locator EXACTLY as printed for that cell — in a merged or grid-shifted row the printed cell ids don't line up with logical column numbers, so don't derive one by counting columns.",
	);
}

/** Parse an `--at`-shaped string for a table-scoped verb. Returns the fully
 * qualified block id of the addressed table — `t0` for top-level, or a chained
 * id like `t0:r0c1:t0` for a table nested inside an outer cell. Cell-chained
 * locators (any depth) are unwrapped recursively so every locator-taking surface
 * accepts the same syntax. Null when the input doesn't address a table. */
export function parseTableAt(at: string): string | null {
	return parseSafely(at, (locator) => composeChainedBlockId(locator, /^t\d+$/));
}

/** Parse an `--at`-shaped string for a row-scoped verb. Returns the containing
 * table's block id (top-level or nested) plus the 0-based row index. */
export function parseRowAt(
	at: string,
): { tableId: string; row: number } | null {
	return parseSafely(at, composeRowTarget);
}

/** Parse an `--at`-shaped string for a column-scoped verb. */
export function parseColumnAt(
	at: string,
): { tableId: string; col: number } | null {
	return parseSafely(at, composeColumnTarget);
}

/** Parse an `--at`-shaped string for a cell-range verb (`merge`). */
export function parseCellRangeAt(at: string): {
	tableId: string;
	start: { row: number; col: number };
	end: { row: number; col: number };
} | null {
	return parseSafely(at, composeCellRangeTarget);
}

/** Parse an `--at`-shaped string for a single-cell verb (`unmerge`). The
 * locator's innermost segment must be a bare cell (no further inner). */
export function parseCellAt(
	at: string,
): { tableId: string; row: number; col: number } | null {
	return parseSafely(at, composeCellTarget);
}

function parseSafely<T>(
	at: string,
	compose: (locator: Locator) => T | null,
): T | null {
	try {
		return compose(parseLocator(at));
	} catch (error) {
		if (!(error instanceof LocatorParseError)) throw error;
		return null;
	}
}

function composeChainedBlockId(
	locator: Locator,
	leafRe: RegExp,
): string | null {
	if (locator.kind === "block") {
		return leafRe.test(locator.blockId) ? locator.blockId : null;
	}
	if (locator.kind === "cell" && locator.inner) {
		const innerId = composeChainedBlockId(locator.inner, leafRe);
		if (innerId === null) return null;
		return `${locator.tableId}:r${locator.row}c${locator.col}:${innerId}`;
	}
	return null;
}

function composeRowTarget(
	locator: Locator,
): { tableId: string; row: number } | null {
	if (locator.kind === "tableRow") {
		return { tableId: locator.tableId, row: locator.row };
	}
	if (locator.kind === "cell" && locator.inner) {
		const inner = composeRowTarget(locator.inner);
		if (!inner) return null;
		return {
			tableId: `${locator.tableId}:r${locator.row}c${locator.col}:${inner.tableId}`,
			row: inner.row,
		};
	}
	return null;
}

function composeColumnTarget(
	locator: Locator,
): { tableId: string; col: number } | null {
	if (locator.kind === "tableColumn") {
		return { tableId: locator.tableId, col: locator.col };
	}
	if (locator.kind === "cell" && locator.inner) {
		const inner = composeColumnTarget(locator.inner);
		if (!inner) return null;
		return {
			tableId: `${locator.tableId}:r${locator.row}c${locator.col}:${inner.tableId}`,
			col: inner.col,
		};
	}
	return null;
}

function composeCellRangeTarget(locator: Locator): {
	tableId: string;
	start: { row: number; col: number };
	end: { row: number; col: number };
} | null {
	if (locator.kind === "cellRange") {
		return {
			tableId: locator.tableId,
			start: locator.start,
			end: locator.end,
		};
	}
	if (locator.kind === "cell" && locator.inner) {
		const inner = composeCellRangeTarget(locator.inner);
		if (!inner) return null;
		return {
			tableId: `${locator.tableId}:r${locator.row}c${locator.col}:${inner.tableId}`,
			start: inner.start,
			end: inner.end,
		};
	}
	return null;
}

function composeCellTarget(
	locator: Locator,
): { tableId: string; row: number; col: number } | null {
	if (locator.kind !== "cell") return null;
	if (!locator.inner) {
		return { tableId: locator.tableId, row: locator.row, col: locator.col };
	}
	const inner = composeCellTarget(locator.inner);
	if (!inner) return null;
	return {
		tableId: `${locator.tableId}:r${locator.row}c${locator.col}:${inner.tableId}`,
		row: inner.row,
		col: inner.col,
	};
}
