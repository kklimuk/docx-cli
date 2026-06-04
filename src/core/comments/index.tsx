import type { Document } from "../ast/document";
import type { LocatorResolveError as LocatorResolveErrorType } from "../locators";
import { LocatorResolveError } from "../locators";
import { XmlNode } from "../parser";
import {
	type AuditCommentAnchor,
	addCommentMarkersAroundRun,
	addCommentMarkersToParagraph,
	addCommentRangeMarkers,
	addReplyCommentMarkers,
	authorInitials,
	CommentBody,
	type CommentSpan,
	generateParaId,
	removeCommentMarkers,
	SpanOutOfRangeError,
} from "./markers";

export {
	type AuditCommentAnchor,
	addCommentMarkersAroundRun,
	addCommentMarkersToParagraph,
	addCommentRangeMarkers,
	addReplyCommentMarkers,
	authorInitials,
	CommentBody,
	type CommentBodyOptions,
	type CommentSpan,
	findContainingParagraph,
	findElementOffsetsInParagraph,
	generateParaId,
	paragraphTextLength,
	removeCommentMarkers,
	SpanOutOfRangeError,
} from "./markers";

import type { Comment } from "../ast/types";
import type { FindView } from "../find";

/** Cross-cutting lens over the document's comments: spans the body (marker
 * placement / removal) and the comments part (mint + persist `<w:comment>`,
 * `<w15:commentEx>`). Constructed at call sites with `new Comments(document)`;
 * holds only a back-reference. Pure-comments-part operations (lookup by id,
 * paraId allocation, `nextId`) live on the embedded `CommentsView`
 * (`document.comments`); this lens orchestrates everything that crosses parts
 * or needs Document-level provisioning like `ensureCommentsExtended`. */
export class Comments {
	constructor(private document: Document) {}

	/** List comments in document order. Filters resolved by default; pass
	 * `includeResolved: true` to keep them. `thread: cN` restricts the result
	 * to that root and its descendants. */
	list(
		options: { includeResolved?: boolean; thread?: string } = {},
	): Comment[] {
		let comments = this.document.body.comments;
		if (!options.includeResolved) {
			comments = comments.filter((comment) => !comment.resolved);
		}
		if (options.thread) {
			const allowed = collectThread(comments, options.thread);
			comments = comments.filter((comment) => allowed.has(comment.id));
		}
		return comments;
	}

	/** Anchor a new comment to a paragraph (optionally a span within it) or
	 * a cross-paragraph range. The CLI converts `--range` / `--anchor` /
	 * batch JSONL into a `CommentAnchorSpec`; this method does the marker
	 * placement, mints a numeric id, and appends the `<w:comment>` body.
	 * Returns the minted numeric id (no `c` prefix).
	 *
	 * Throws `CommentsError("BLOCK_NOT_FOUND")` on a stale blockId, or
	 * `CommentsError("INVALID_LOCATOR")` when the span exceeds the paragraph
	 * text length under the requested `findView`. */
	add(
		anchor: CommentAnchorSpec,
		options: {
			body: string;
			author: string;
			date?: string;
			findView?: FindView;
		},
	): string {
		const date = options.date ?? new Date().toISOString();
		const numericId = this.document.comments?.nextId() ?? "0";
		const paraId = generateParaId();

		try {
			if (anchor.kind === "single") {
				const ref = this.#resolveBlock(anchor.blockId);
				addCommentMarkersToParagraph(
					ref.node,
					numericId,
					anchor.span,
					options.findView,
				);
			} else {
				const startRef = this.#resolveBlock(anchor.startBlockId);
				const endRef = this.#resolveBlock(anchor.endBlockId);
				addCommentRangeMarkers(
					startRef.node,
					anchor.startOffset,
					endRef.node,
					anchor.endOffset,
					numericId,
					options.findView,
				);
			}
		} catch (error) {
			if (error instanceof SpanOutOfRangeError) {
				throw new CommentsError("INVALID_LOCATOR", error.message);
			}
			throw error;
		}

		this.#appendCommentBody(
			numericId,
			paraId,
			options.body,
			options.author,
			date,
		);
		return numericId;
	}

	/** Append a reply chained to `parentId` via `<w15:commentEx>`. Mints +
	 * persists a `w14:paraId` on the parent first if absent. Returns the
	 * minted numeric id (no `c` prefix). Throws
	 * `CommentsError("COMMENT_NOT_FOUND")` if the parent doesn't exist. */
	reply(
		parentId: string,
		body: string,
		options: { author: string; date?: string },
	): string {
		const parentNumericId = parentId.startsWith("c")
			? parentId.slice(1)
			: parentId;
		const view = this.document.comments;
		if (!view?.findById(parentNumericId)) {
			throw new CommentsError(
				"COMMENT_NOT_FOUND",
				`Parent comment not found: c${parentNumericId}`,
			);
		}
		const parentParaId = view.ensureParaId(parentNumericId);
		if (!parentParaId) {
			throw new CommentsError(
				"COMMENT_NOT_FOUND",
				`Parent comment c${parentNumericId} could not be assigned a w14:paraId.`,
			);
		}

		const date = options.date ?? new Date().toISOString();
		const numericId = view.nextId();
		const replyParaId = generateParaId();

		// Word drops any comment without a `<w:commentReference>` in the body;
		// mirror the parent thread's anchor markers so the reply renders. Anchor
		// before writing any part so an unanchorable parent aborts cleanly.
		const anchored = addReplyCommentMarkers(
			this.document.documentTree,
			parentNumericId,
			numericId,
		);
		if (!anchored) {
			throw new CommentsError(
				"COMMENT_NOT_FOUND",
				`Parent comment c${parentNumericId} has no anchor in the document body; cannot anchor reply.`,
			);
		}

		this.#appendCommentBody(
			numericId,
			replyParaId,
			body,
			options.author,
			date,
			parentParaId,
			parentNumericId,
		);

		const extView = this.document.ensureCommentsExtended();
		const extRoot = extView.extendedTree
			? XmlNode.findRoot(extView.extendedTree, "w15:commentsEx")
			: undefined;
		if (!extRoot) throw new Error("expected <w15:commentsEx> root");

		let parentEntry = extRoot.children.find(
			(child) =>
				child.tag === "w15:commentEx" &&
				child.getAttribute("w15:paraId") === parentParaId,
		);
		if (!parentEntry) {
			parentEntry = new XmlNode("w15:commentEx", {
				"w15:paraId": parentParaId,
				"w15:done": "0",
			});
			extRoot.children.push(parentEntry);
		}

		const replyEntry = new XmlNode("w15:commentEx", {
			"w15:paraId": replyParaId,
			"w15:paraIdParent": parentParaId,
			"w15:done": "0",
		});
		const parentEntryIndex = extRoot.children.indexOf(parentEntry);
		if (parentEntryIndex !== -1) {
			// Scan forward from the parent to find the last existing reply
			// in this thread so the new reply appends after the tail.
			let insertAfterIndex = parentEntryIndex;
			for (let i = parentEntryIndex + 1; i < extRoot.children.length; i++) {
				const child = extRoot.children[i];
				if (
					child.tag === "w15:commentEx" &&
					child.getAttribute("w15:paraIdParent") === parentParaId
				) {
					insertAfterIndex = i;
				}
			}
			extRoot.children.splice(insertAfterIndex + 1, 0, replyEntry);
		} else {
			extRoot.children.push(replyEntry);
		}

		return numericId;
	}

	/** Mark every id in `ids` as resolved (or, with `resolved: false`,
	 * unresolved). The whole batch is pre-validated against the current
	 * tree so the apply loop is atomic — any unknown id aborts before any
	 * mutation lands. Throws `CommentsError("COMMENT_NOT_FOUND")` with the
	 * first offending id. */
	resolve(ids: string[], resolved: boolean): void {
		const normalized = ids.map((id) => (id.startsWith("c") ? id : `c${id}`));
		const view = this.document.comments;
		const paraIdByCommentId = new Map<string, string>();
		for (const commentId of normalized) {
			const numericId = commentId.slice(1);
			if (!view?.findById(numericId)) {
				throw new CommentsError(
					"COMMENT_NOT_FOUND",
					`Comment not found: ${commentId}`,
				);
			}
			const paraId = view.ensureParaId(numericId);
			if (!paraId) {
				throw new CommentsError(
					"COMMENT_NOT_FOUND",
					`Comment ${commentId} could not be assigned a w14:paraId.`,
				);
			}
			paraIdByCommentId.set(commentId, paraId);
		}

		const extView = this.document.ensureCommentsExtended();
		const extRoot = extView.extendedTree
			? XmlNode.findRoot(extView.extendedTree, "w15:commentsEx")
			: undefined;
		if (!extRoot) throw new Error("expected <w15:commentsEx> root");

		for (const commentId of normalized) {
			const paraId = paraIdByCommentId.get(commentId);
			if (!paraId) continue;
			let entry = extRoot.children.find(
				(child) =>
					child.tag === "w15:commentEx" &&
					child.getAttribute("w15:paraId") === paraId,
			);
			if (!entry) {
				entry = new XmlNode("w15:commentEx", { "w15:paraId": paraId });
				extRoot.children.push(entry);
			}
			if (resolved) entry.setAttribute("w15:done", "1");
			else delete entry.attributes["w15:done"];
		}
	}

	/** Remove every id in `ids`: splice the `<w:comment>` body, prune any
	 * `<w15:commentEx>` entry keyed to its paraId, and strip
	 * `<w:commentRangeStart>` / `<w:commentRangeEnd>` /
	 * `<w:commentReference>` markers from the body. Deleting a thread parent
	 * cascades through its replies so none are left with dangling markers or a
	 * `w15:paraIdParent` pointing at a removed comment. Pre-validates the
	 * batch — any unknown id aborts before any mutation. Throws
	 * `CommentsError("COMMENT_NOT_FOUND")` with the first offending id. */
	delete(ids: string[]): void {
		const normalized = ids.map((id) => (id.startsWith("c") ? id : `c${id}`));
		const view = this.document.comments;
		for (const commentId of normalized) {
			const numericId = commentId.slice(1);
			if (!view?.findById(numericId)) {
				throw new CommentsError(
					"COMMENT_NOT_FOUND",
					`Comment not found: ${commentId}`,
				);
			}
		}
		if (!view) return; // `normalized` is empty (no ids → no `view` access yet).

		const expanded = new Set(normalized);
		for (const commentId of normalized) {
			for (const replyId of view.descendantReplyIds(commentId)) {
				expanded.add(`c${replyId}`);
			}
		}

		const root = XmlNode.findRoot(view.tree, "w:comments");
		if (!root) return;

		for (const commentId of expanded) {
			const numericId = commentId.slice(1);
			const node = view.findById(numericId);
			if (!node) continue; // pre-validated above
			const paraId = view.paraIdFor(numericId);

			const index = root.children.indexOf(node);
			if (index !== -1) root.children.splice(index, 1);

			if (paraId && view.extendedTree) {
				const extRoot = XmlNode.findRoot(view.extendedTree, "w15:commentsEx");
				if (extRoot) {
					extRoot.children = extRoot.children.filter(
						(child) =>
							!(
								child.tag === "w15:commentEx" &&
								child.getAttribute("w15:paraId") === paraId
							),
					);
				}
			}

			removeCommentMarkers(this.document.documentTree, numericId);
		}
	}

	/** Emit a comment for operations OOXML can't track natively (hyperlink
	 * edits, image replacement). Anchored either to a text span (default) or
	 * around a specific run (image runs, where text length is zero).
	 * Materializes the comments part. Returns the minted numeric id. */
	addAudit(
		anchor: AuditCommentAnchor,
		options: { body: string; author: string; date: string },
	): string {
		const numericId = this.document.comments?.nextId() ?? "0";
		const paraId = generateParaId();

		if (anchor.kind === "span") {
			addCommentMarkersToParagraph(anchor.paragraph, numericId, anchor.span);
		} else {
			addCommentMarkersAroundRun(anchor.paragraph, anchor.run, numericId);
		}

		this.#appendCommentBody(
			numericId,
			paraId,
			options.body,
			options.author,
			options.date,
		);
		return numericId;
	}

	#appendCommentBody(
		numericId: string,
		paraId: string,
		text: string,
		author: string,
		date: string,
		paraIdParent?: string,
		afterCommentId?: string,
	): void {
		const commentsView = this.document.ensureComments();
		const commentsRoot = XmlNode.findRoot(commentsView.tree, "w:comments");
		if (!commentsRoot) throw new Error("expected <w:comments> root");
		const body = (
			<CommentBody
				options={{
					id: numericId,
					author,
					date,
					initials: authorInitials(author),
					paraId,
					...(paraIdParent ? { paraIdParent } : {}),
					text,
				}}
			/>
		);
		if (afterCommentId) {
			commentsView.insertReplyAfter(afterCommentId, body);
		} else {
			commentsRoot.children.push(body);
		}
	}

	#resolveBlock(blockId: string): { node: XmlNode; parent: XmlNode[] } {
		try {
			return this.document.body.resolveBlock(blockId);
		} catch (error) {
			if (error instanceof LocatorResolveError) {
				throw new CommentsError(
					"BLOCK_NOT_FOUND",
					(error as LocatorResolveErrorType).message,
				);
			}
			throw error;
		}
	}
}

/** Anchor for `Comments.add`. The CLI parses `--range LOCATOR` and resolves
 * `--anchor PHRASE` (via the find subsystem) into one of these two shapes:
 * single block (optionally with a sub-paragraph span), or cross-paragraph
 * range with explicit start/end offsets. */
export type CommentAnchorSpec =
	| { kind: "single"; blockId: string; span?: CommentSpan }
	| {
			kind: "range";
			startBlockId: string;
			startOffset: number;
			endBlockId: string;
			endOffset: number;
	  };

/** Domain error from `Comments.*`. `code` is a literal subset of the CLI's
 * `ErrorCode` union so callers can `return fail(err.code, err.message,
 * err.hint)` directly — no cast, full type-check coverage. */
export type CommentsErrorCode =
	| "COMMENT_NOT_FOUND"
	| "BLOCK_NOT_FOUND"
	| "INVALID_LOCATOR";

export class CommentsError extends Error {
	constructor(
		public code: CommentsErrorCode,
		message: string,
		public hint?: string,
	) {
		super(message);
		this.name = "CommentsError";
	}
}

function collectThread(comments: Comment[], rootId: string): Set<string> {
	const allowed = new Set<string>([rootId]);
	let changed = true;
	while (changed) {
		changed = false;
		for (const comment of comments) {
			if (allowed.has(comment.id)) continue;
			if (comment.parentId && allowed.has(comment.parentId)) {
				allowed.add(comment.id);
				changed = true;
			}
		}
	}
	return allowed;
}
