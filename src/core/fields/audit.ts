import type { PackagePart } from "../raw/package-consistency";
import { parseFields } from "./parse";

/** `validate`'s field-balance check: every `<w:fldChar>` the begin/separate/
 *  end stack can't pair, per story part. A part that holds several stories
 *  (footnotes, endnotes, comments) is parsed one story per note/comment — a
 *  field never spans two. Nothing is fixable: which half to drop, or where the
 *  missing end belongs, isn't knowable from the XML. */
export function auditFieldBalance(parts: PackagePart[]): FieldIssue[] {
	const issues: FieldIssue[] = [];
	for (const part of parts) {
		const stories = MULTI_STORY_ROOTS.has(part.root.tag)
			? part.root.children.filter((child) => !child.isText)
			: [part.root];
		const { strays } = parseFields(stories, "current");
		if (strays.length === 0) continue;
		const counts = new Map<string, number>();
		for (const stray of strays) {
			const type = stray.node.getAttribute("w:fldCharType") ?? "?";
			counts.set(type, (counts.get(type) ?? 0) + 1);
		}
		const described = [...counts]
			.map(([type, count]) => `${count} "${type}"`)
			.join(", ");
		issues.push({
			part: part.name,
			message: `unbalanced field: ${described} fldChar marker${strays.length === 1 ? "" : "s"} with no matching begin/end — half a field (Word may show the instruction as text, or swallow the text after an unclosed field)`,
			severity: "error",
			fixable: false,
			hint: "Find the paragraph with `docx raw get FILE --at pN` (look for w:fldChar) and repair it with `docx raw edit`, or re-insert the field in Word.",
		});
	}
	return issues;
}

const MULTI_STORY_ROOTS = new Set(["w:footnotes", "w:endnotes", "w:comments"]);

type FieldIssue = {
	part: string;
	message: string;
	severity: "error";
	fixable: false;
	hint: string;
};
