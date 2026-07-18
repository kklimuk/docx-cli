import type { Document } from "@core";
import type { XmlNode } from "@core/parser";
import { addAuditComment } from "../audit-comment";
import { resolveTracked } from "../respond";

/** Record a cell merge / unmerge under track-changes. Word applies table-cell
 * merges immediately even with tracking on — it does NOT emit a revision marker
 * (verified empirically: a Word merge with track-changes on produces a plain
 * `<w:gridSpan>` with no `<w:cellMerge>`/`<w:tcPrChange>`). So we match Word:
 * apply the merge in place and anchor a `[docx-cli]` audit comment (the shared
 * convention in `cli/audit-comment.ts`) so the structural change is still
 * visible in review. No-op when tracking is off. */
export function noteStructuralChange(
	document: Document,
	anchorCell: XmlNode | undefined,
	message: string,
	authorFlag: string | undefined,
	trackFlag?: boolean,
): void {
	if (!resolveTracked(document, trackFlag)) return;
	const paragraph = anchorCell?.findChild("w:p");
	if (!paragraph) return;
	addAuditComment(
		document,
		{ kind: "span", paragraph, span: { start: 0, end: 0 } },
		message,
		authorFlag,
	);
}
