import type { BlockReference, DocView } from "../ast";
import type { XmlNode } from "../parser";
import type { Locator } from "./parse";

export class LocatorResolveError extends Error {
	constructor(
		public locator: Locator,
		message: string,
	) {
		super(message);
		this.name = "LocatorResolveError";
	}
}

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

export function resolveBlock(view: DocView, blockId: string): BlockReference {
	const reference = view.blockReferences.get(blockId);
	if (!reference) {
		throw new LocatorResolveError(
			{ kind: "block", blockId },
			`Block not found: ${blockId}`,
		);
	}
	return reference;
}

/** A contiguous span of blocks resolved from a `pN-pM` locator. The two
 * endpoints MUST share a parent (both top-level body children, or both
 * inside the same cell) — cross-parent ranges aren't meaningful. */
export type BlockRangeReference = {
	parent: XmlNode[];
	startIndex: number;
	endIndex: number;
};

/** Resolve a `pN-pM` block-range locator. Validates that the two endpoints
 * live under the same parent array (so they can be spliced as a unit) and
 * that the start index is ≤ the end index in document order. */
export function resolveBlockRange(
	view: DocView,
	startBlockId: string,
	endBlockId: string,
): BlockRangeReference {
	const start = resolveBlock(view, startBlockId);
	const end = resolveBlock(view, endBlockId);
	if (start.parent !== end.parent) {
		throw new LocatorResolveError(
			{ kind: "blockRange", startBlockId, endBlockId },
			`Range endpoints ${startBlockId} and ${endBlockId} live in different containers — they must be siblings`,
		);
	}
	const startIndex = start.parent.indexOf(start.node);
	const endIndex = end.parent.indexOf(end.node);
	if (startIndex === -1 || endIndex === -1) {
		throw new LocatorResolveError(
			{ kind: "blockRange", startBlockId, endBlockId },
			"Range endpoint became detached from its parent (stale block reference)",
		);
	}
	if (endIndex < startIndex) {
		throw new LocatorResolveError(
			{ kind: "blockRange", startBlockId, endBlockId },
			`Range ${startBlockId}-${endBlockId} runs backwards — ${endBlockId} appears before ${startBlockId} in document order`,
		);
	}
	return { parent: start.parent, startIndex, endIndex };
}
