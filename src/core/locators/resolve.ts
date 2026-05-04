import type {
	BlockReference,
	CommentReference,
	DocView,
	ImageReference,
} from "../ast";
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

export function resolveComment(
	view: DocView,
	commentId: string,
): CommentReference {
	const reference = view.commentReferences.get(commentId);
	if (!reference) {
		throw new LocatorResolveError(
			{ kind: "comment", commentId },
			`Comment not found: ${commentId}`,
		);
	}
	return reference;
}

export function resolveImage(view: DocView, imageId: string): ImageReference {
	const reference = view.imageById.get(imageId);
	if (!reference) {
		throw new LocatorResolveError(
			{ kind: "image", imageId },
			`Image not found: ${imageId}`,
		);
	}
	return reference;
}
