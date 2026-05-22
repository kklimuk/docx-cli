import {
	createRevisionAllocator,
	type DocView,
	isTrackChangesEnabled,
	resolveAuthor,
	resolveDate,
	type TrackedMeta,
} from "@core";
import type { XmlNode } from "@core/parser";
import { emitAuditComment } from "../comments/helpers";

/** A fresh tracked-change meta (author/date/revisionId) for the current view. */
export function mintRevisionMeta(
	view: DocView,
	authorFlag: string | undefined,
): TrackedMeta {
	return {
		author: resolveAuthor(authorFlag),
		date: resolveDate(),
		revisionId: createRevisionAllocator(view).next(),
	};
}

/** Resolve a `tN` locator to its `<w:tbl>` element, or null when the id is
 * absent or not a table. */
export function resolveTableNode(
	view: DocView,
	tableId: string,
): XmlNode | null {
	const reference = view.blockReferences.get(tableId);
	if (!reference || reference.node.tag !== "w:tbl") return null;
	return reference.node;
}

/** Record a cell merge / unmerge under track-changes. Word applies table-cell
 * merges immediately even with tracking on — it does NOT emit a revision marker
 * (verified empirically: a Word merge with track-changes on produces a plain
 * `<w:gridSpan>` with no `<w:cellMerge>`/`<w:tcPrChange>`). So we match Word:
 * apply the merge in place and, mirroring the hyperlinks/images audit-comment
 * pattern, anchor a `[docx-cli]` comment so the structural change is still
 * visible in review. No-op when tracking is off. */
export function noteStructuralChange(
	view: DocView,
	anchorCell: XmlNode | undefined,
	message: string,
	authorFlag: string | undefined,
): void {
	if (!isTrackChangesEnabled(view)) return;
	const paragraph = anchorCell?.findChild("w:p");
	if (!paragraph) return;
	emitAuditComment(
		view,
		{ kind: "span", paragraph, span: { start: 0, end: 0 } },
		{
			body: `[docx-cli] ${message}`,
			author: resolveAuthor(authorFlag),
			date: resolveDate(),
		},
	);
}
