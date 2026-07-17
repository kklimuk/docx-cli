import { type Document, resolveAuthor, resolveDate } from "@core";
import { type AuditCommentAnchor, Comments } from "@core/comments";

/** The `[docx-cli]` audit-comment convention in one place: when a CLI change
 *  has no honest tracked-change construct (table merges/borders, raw XML,
 *  section-break removal, image swaps), the mutation applies untracked and a
 *  comment marks the spot for review. Callers own their gate (tables check
 *  `resolveTracked`, raw checks the doc toggle only) and their `anchor` choice
 *  (a zero-width span, or wrapped around a specific run); this owns the
 *  `[docx-cli]` prefix and the author/date chain. Lives in `cli/` because that
 *  prefix is CLI policy, not a model concern. */
export function addAuditComment(
	document: Document,
	anchor: AuditCommentAnchor,
	message: string,
	authorFlag: string | undefined,
): void {
	new Comments(document).addAudit(anchor, {
		body: `[docx-cli] ${message}`,
		author: resolveAuthor(authorFlag),
		date: resolveDate(),
	});
}
