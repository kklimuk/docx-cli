import type { Document } from "../ast/document";
import { XmlNode } from "../parser";
import {
	type AuditCommentAnchor,
	addCommentMarkersAroundRun,
	addCommentMarkersToParagraph,
	authorInitials,
	CommentBody,
	generateParaId,
} from "./markers";

export {
	type AuditCommentAnchor,
	addCommentMarkersAroundRun,
	addCommentMarkersToParagraph,
	addCommentRangeMarkers,
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

/** Cross-cutting lens over the document's comments part: minting audit
 * comments, looking comments up by id, and resolving the `w14:paraId` that
 * threads/resolution key off of. Constructed at call sites with
 * `new Comments(document)`; holds only a back-reference. The CLI's comment
 * verbs (add/reply/resolve/delete) orchestrate the marker functions in
 * `markers.tsx` directly; this lens carries the operations that span the
 * document body + the comments part. */
export class Comments {
	constructor(private document: Document) {}

	/** Find a comment by its numeric id (no `c` prefix) — returns the
	 * `<w:comment>` node + its parent array for splicing, or undefined. */
	findById(
		numericId: string,
	): { node: XmlNode; parent: XmlNode[] } | undefined {
		if (!this.document.comments?.tree) return undefined;
		const root = XmlNode.findRoot(this.document.comments.tree, "w:comments");
		if (!root) return undefined;
		for (const child of root.children) {
			if (
				child.tag === "w:comment" &&
				child.getAttribute("w:id") === numericId
			) {
				return { node: child, parent: root.children };
			}
		}
		return undefined;
	}

	/** Read a comment's `w14:paraId`, or undefined if the comment or its inner
	 * paragraph is missing. Accepts the `cN` or bare-`N` id form. */
	paraIdFor(commentId: string): string | undefined {
		const paragraph = this.commentParagraph(commentId);
		return paragraph?.getAttribute("w14:paraId");
	}

	/** Read a comment's `w14:paraId`, minting + persisting one (and the
	 * `xmlns:w14` declaration) if absent. reply/resolve key off paraId, so
	 * this guarantees one exists. Accepts the `cN` or bare-`N` id form. */
	ensureParaId(commentId: string): string | undefined {
		const root = this.document.comments?.tree
			? XmlNode.findRoot(this.document.comments.tree, "w:comments")
			: undefined;
		const paragraph = this.commentParagraph(commentId);
		if (!root || !paragraph) return undefined;
		const existing = paragraph.getAttribute("w14:paraId");
		if (existing) return existing;
		const fresh = generateParaId();
		paragraph.setAttribute("w14:paraId", fresh);
		if (!root.attributes["xmlns:w14"]) {
			root.setAttribute(
				"xmlns:w14",
				"http://schemas.microsoft.com/office/word/2010/wordml",
			);
		}
		return fresh;
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

		const commentsView = this.document.ensureComments();
		const commentsRoot = XmlNode.findRoot(commentsView.tree, "w:comments");
		if (!commentsRoot) throw new Error("expected <w:comments> root");
		commentsRoot.children.push(
			<CommentBody
				options={{
					id: numericId,
					author: options.author,
					date: options.date,
					initials: authorInitials(options.author),
					paraId,
					text: options.body,
				}}
			/>,
		);

		return numericId;
	}

	private commentParagraph(commentId: string): XmlNode | undefined {
		if (!this.document.comments?.tree) return undefined;
		const root = XmlNode.findRoot(this.document.comments.tree, "w:comments");
		if (!root) return undefined;
		const numericId = commentId.startsWith("c")
			? commentId.slice(1)
			: commentId;
		for (const child of root.children) {
			if (child.tag !== "w:comment") continue;
			if (child.getAttribute("w:id") !== numericId) continue;
			return child.findChild("w:p");
		}
		return undefined;
	}
}
