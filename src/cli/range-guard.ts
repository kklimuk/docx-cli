import type { BlockRangeReference } from "@core";
import { fail } from "./respond";

/** Reject a tracked range edit/delete whose `pN-pM` span includes a non-
 *  paragraph block (most commonly a table). The tracked-range helpers in
 *  `@core/track-changes/replace` walk every block in `[startIndex, endIndex]`
 *  and call `markParagraphMarkAs`, which unconditionally injects a `<w:pPr>` —
 *  doing that to a `<w:tbl>` produces invalid OOXML ("unreadable content").
 *
 *  We don't try to track-delete a table here. Agents should either:
 *  - Toggle tracking off, then range-edit/delete spans that include tables, or
 *  - Delete the table separately via `docx delete --at tN` (untracked) and
 *    handle the surrounding paragraphs with their own tracked operations.
 *
 *  Untracked range paths splice through tables cleanly, so this guard only
 *  fires when tracking is on. */
export async function rejectNonParagraphTrackedRange(
	rangeRef: BlockRangeReference,
	locator: string,
): Promise<number | null> {
	for (let i = rangeRef.startIndex; i <= rangeRef.endIndex; i++) {
		const block = rangeRef.parent[i];
		if (block && block.tag !== "w:p") {
			const what = block.tag === "w:tbl" ? "a table" : `a ${block.tag} block`;
			return fail(
				"TRACKED_CHANGE_CONFLICT",
				`Range ${locator} spans ${what} — tracked range edits/deletes don't support that.`,
				"Toggle tracking off (`docx track-changes off`), or handle the table separately via `docx delete --at tN`.",
			);
		}
	}
	return null;
}
