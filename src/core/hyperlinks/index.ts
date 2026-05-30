import type { Document } from "../ast/document";
import {
	Comments,
	findContainingParagraph,
	findElementOffsetsInParagraph,
} from "../comments";
import type { XmlNode } from "../parser";
import { resolveAuthor, resolveDate } from "../track-changes";
import { type Span, wrapSpanInHyperlink } from "./wrap";

export { HyperlinkWrapError, type Span, wrapSpanInHyperlink } from "./wrap";

/** Cross-cutting lens over the document's hyperlinks. Constructed at call
 * sites with `new Hyperlinks(document)`; holds only a back-reference. A
 * hyperlink owns a relationship (not its text), so these operations thread
 * the document body (the `<w:hyperlink>` wrapper) and the relationships part
 * together. OOXML has no `<w:hyperlinkChange>` element, so under tracking
 * each mutation drops a `[docx-cli]` audit comment instead of a real
 * tracked-change wrapper (via `Comments.addAudit`). */
export class Hyperlinks {
	constructor(private document: Document) {}

	/** Wrap a span's runs in a `<w:hyperlink>`: mint the relationship, split
	 * the runs at the span edges, record the rel→url map, and (under
	 * tracking) drop an audit comment. Throws `HyperlinkWrapError` if the span
	 * overlaps an existing wrapper. */
	add(
		paragraph: XmlNode,
		span: Span,
		url: string,
		options: { author?: string } = {},
	): void {
		const relationshipId = this.document.relationships.addHyperlink(url);
		wrapSpanInHyperlink(paragraph, span, relationshipId);
		this.document.relationships.hyperlinksByRelationshipId.set(relationshipId, {
			url,
		});
		if (this.document.isTrackChangesEnabled()) {
			new Comments(this.document).addAudit(
				{ kind: "span", paragraph, span },
				auditBody(`hyperlink added → ${url}`, options.author),
			);
		}
	}

	/** Repoint a hyperlink at a new URL. If its relationship is shared by
	 * multiple `<w:hyperlink>` elements, mints a fresh relationship so the
	 * others are unaffected; otherwise rewrites the existing target in place.
	 * Returns the prior URL. Throws `HyperlinkNotFoundError`. */
	replace(
		id: string,
		url: string,
		options: { author?: string } = {},
	): { from?: string } {
		const reference = this.document.body.hyperlinkById.get(id);
		if (!reference) throw new HyperlinkNotFoundError(id);

		const existingId = reference.relationshipId;
		const from = existingId
			? this.document.relationships.hyperlinksByRelationshipId.get(existingId)
					?.url
			: undefined;
		const shared = existingId
			? countHyperlinkUsages(this.document.documentTree, existingId) > 1
			: true;

		if (!existingId || shared) {
			const newId = this.document.relationships.addHyperlink(url);
			reference.node.setAttribute("r:id", newId);
			reference.relationshipId = newId;
		} else {
			this.document.relationships.setHyperlinkTarget(existingId, url);
		}

		if (this.document.isTrackChangesEnabled()) {
			this.auditAtNode(
				reference.node,
				`hyperlink target changed: ${from ?? "(none)"} → ${url}`,
				options.author,
			);
		}
		return { from };
	}

	/** Unwrap a hyperlink (its display text stays), prune the relationship if
	 * nothing else references it, and (under tracking) drop an audit comment.
	 * Returns the prior URL. Throws `HyperlinkNotFoundError`, or
	 * `HyperlinkStaleError` if the node detached from its parent. */
	delete(id: string, options: { author?: string } = {}): { from?: string } {
		const reference = this.document.body.hyperlinkById.get(id);
		if (!reference) throw new HyperlinkNotFoundError(id);

		const from = reference.relationshipId
			? this.document.relationships.hyperlinksByRelationshipId.get(
					reference.relationshipId,
				)?.url
			: undefined;

		const index = reference.parent.indexOf(reference.node);
		if (index === -1) throw new HyperlinkStaleError(id);

		// Capture the anchor span BEFORE unwrapping — once the wrapper is gone
		// the node can't be located in its paragraph.
		const trackingOn = this.document.isTrackChangesEnabled();
		const paragraph = trackingOn
			? findContainingParagraph(this.document.documentTree, reference.node)
			: null;
		const offsets =
			trackingOn && paragraph
				? findElementOffsetsInParagraph(paragraph, reference.node)
				: null;

		reference.parent.splice(index, 1, ...reference.node.children);
		this.document.body.hyperlinkById.delete(id);

		// Prune only when the rId is referenced nowhere — scanning every
		// attribute, not just `<w:hyperlink>`, so we don't dangle an
		// `<a:hlinkClick r:id>` in a drawing or a `<w:fldSimple>` HYPERLINK
		// field that shares this relationship (root CLAUDE.md invariant).
		if (reference.relationshipId) {
			this.document.relationships.removeIfUnreferenced(
				reference.relationshipId,
				this.document.documentTree,
			);
		}

		if (trackingOn && paragraph && offsets) {
			new Comments(this.document).addAudit(
				{ kind: "span", paragraph, span: offsets },
				auditBody(
					`hyperlink removed (was: ${from ?? "(none)"})`,
					options.author,
				),
			);
		}
		return { from };
	}

	private auditAtNode(
		node: XmlNode,
		message: string,
		author: string | undefined,
	): void {
		const paragraph = findContainingParagraph(this.document.documentTree, node);
		const offsets = paragraph
			? findElementOffsetsInParagraph(paragraph, node)
			: null;
		if (!paragraph || !offsets) return;
		new Comments(this.document).addAudit(
			{ kind: "span", paragraph, span: offsets },
			auditBody(message, author),
		);
	}
}

/** Count how many `<w:hyperlink>` elements in `documentTree` reference `rId`.
 * The shared-relationship gate for `replace` (a shared rId must fork so the
 * other hyperlinks keep their old target) and a useful dry-run preview. */
export function countHyperlinkUsages(
	documentTree: XmlNode[],
	relationshipId: string,
): number {
	let count = 0;
	function walk(node: XmlNode): void {
		if (
			node.tag === "w:hyperlink" &&
			node.getAttribute("r:id") === relationshipId
		) {
			count += 1;
		}
		for (const child of node.children) walk(child);
	}
	for (const root of documentTree) walk(root);
	return count;
}

function auditBody(
	message: string,
	author: string | undefined,
): { body: string; author: string; date: string } {
	return {
		body: `[docx-cli] ${message}`,
		author: resolveAuthor(author),
		date: resolveDate(),
	};
}

/** Thrown by `Hyperlinks.replace` / `delete` when the id doesn't resolve. */
export class HyperlinkNotFoundError extends Error {
	constructor(public id: string) {
		super(`Hyperlink not found: ${id}`);
		this.name = "HyperlinkNotFoundError";
	}
}

/** Thrown by `Hyperlinks.delete` when the cached node is no longer in its
 * parent (stale reference). */
export class HyperlinkStaleError extends Error {
	constructor(public id: string) {
		super(`Hyperlink reference is stale (parent does not contain it): ${id}`);
		this.name = "HyperlinkStaleError";
	}
}
