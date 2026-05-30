import {
	type BlockRangeReference,
	LocatorResolveError,
} from "../ast/document/body";
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
