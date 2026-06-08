import { type Document, resolveAuthor, resolveDate } from "@core";
import { Comments } from "@core/comments";
import type { XmlNode } from "@core/parser";
import { resolveTracked } from "../respond";

/** Record a cell merge / unmerge under track-changes. Word applies table-cell
 * merges immediately even with tracking on — it does NOT emit a revision marker
 * (verified empirically: a Word merge with track-changes on produces a plain
 * `<w:gridSpan>` with no `<w:cellMerge>`/`<w:tcPrChange>`). So we match Word:
 * apply the merge in place and, mirroring the hyperlinks/images audit-comment
 * pattern, anchor a `[docx-cli]` comment so the structural change is still
 * visible in review. No-op when tracking is off.
 *
 * Lives in `cli/` (not `core/table/`) because the `[docx-cli] …` audit-comment
 * prefix is a CLI policy decision, not a model concern. */
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
	new Comments(document).addAudit(
		{ kind: "span", paragraph, span: { start: 0, end: 0 } },
		{
			body: `[docx-cli] ${message}`,
			author: resolveAuthor(authorFlag),
			date: resolveDate(),
		},
	);
}
