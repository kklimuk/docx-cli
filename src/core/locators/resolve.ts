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
