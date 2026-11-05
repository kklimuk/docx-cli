import { isWrapperVisibleInView } from "../find/replace-span";
import {
	alternateContentBranch,
	collectTextBoxContents,
	isAlternateContent,
} from "../mc";
import { wrapperContent, XmlNode } from "../parser";
import type { RevisionView } from "../track-changes/revision-visibility";

/** Extract a header/footer part's text for the AST `Marginal.text`, with
 *  `<w:fldSimple>` fields rendered as readable tokens (`{page}`, `{date}`,
 *  `{styleref:Heading 1}`, …) rather than their cached value. Paragraphs join
 *  with newlines; an inter-zone `<w:tab/>` becomes a literal tab. Tracked
 *  changes read in `view` — by default the accepted one (insertions shown,
 *  deletions hidden), which is the AST's `Marginal.text`; `read --current`
 *  passes `"current"` plus a `markRevision` that wraps each revision
 *  wrapper's text (the CriticMarkup `read` uses for the body). Pure read
 *  helper (no JSX), shared by the AST reader and `read`. */
export function marginalText(
	tree: XmlNode[],
	view: RevisionView = "accepted",
	markRevision?: (text: string, wrapper: XmlNode) => string,
): string {
	const root =
		XmlNode.findRoot(tree, "w:hdr") ?? XmlNode.findRoot(tree, "w:ftr");
	if (!root) return "";
	const reading: Reading = { view, markRevision };
	const paragraphs: string[] = [];
	for (const child of root.children) {
		if (child.tag !== "w:p") continue;
		paragraphs.push(paragraphText(child, reading));
	}
	return paragraphs.join("\n").trim();
}

type Reading = {
	view: RevisionView;
	markRevision?: (text: string, wrapper: XmlNode) => string;
};

function paragraphText(paragraph: XmlNode, reading: Reading): string {
	let out = "";
	for (const child of paragraph.children) out += nodeText(child, reading);
	return out;
}

function nodeText(node: XmlNode, reading: Reading): string {
	if (node.tag === "w:fldSimple") {
		return fieldToken(node.getAttribute("w:instr") ?? "");
	}
	if (node.tag === "w:r") {
		let out = "";
		for (const child of node.children) {
			// A deleted run's text is `<w:delText>`; it's only reached when the view
			// descends the deletion (`--current`/`--baseline`).
			if (child.tag === "w:t" || child.tag === "w:delText")
				out += child.collectText();
			else if (child.tag === "w:tab") out += "\t";
			else if (child.tag === "w:br" || child.tag === "w:cr") out += "\n";
			// A letterhead text box (Word: `<mc:AlternateContent>` around the
			// shape; legacy: a bare `<w:pict>`): its story's paragraphs join the
			// header text on their own lines, so `headers list` shows what the
			// page shows. Bracketed so the box reads as one unit.
			else if (
				isAlternateContent(child) ||
				child.tag === "w:drawing" ||
				child.tag === "w:pict"
			) {
				out += textBoxText(child, reading);
			}
		}
		return out;
	}
	if (isAlternateContent(node)) {
		const branch = alternateContentBranch(node);
		return branch ? paragraphText(branch, reading) : "";
	}
	// Run-bearing wrappers (a hyperlink, a tracked insertion) — descend into
	// those the view shows; a wrapper it hides contributes nothing.
	if (isWrapperVisibleInView(node.tag, reading.view)) {
		let out = "";
		for (const child of wrapperContent(node)) out += nodeText(child, reading);
		const isRevision = REVISION_WRAPPER_TAGS.has(node.tag);
		return isRevision && out && reading.markRevision
			? reading.markRevision(out, node)
			: out;
	}
	return "";
}

const REVISION_WRAPPER_TAGS = new Set([
	"w:ins",
	"w:del",
	"w:moveFrom",
	"w:moveTo",
]);

function textBoxText(shape: XmlNode, reading: Reading): string {
	const stories = collectTextBoxContents(shape);
	if (stories.length === 0) return "";
	return stories
		.map((story) => {
			const lines = story
				.findChildren("w:p")
				.map((paragraph) => paragraphText(paragraph, reading));
			return `\n[${lines.join("\n")}]`;
		})
		.join("");
}

/** Map a field's `w:instr` to a stable read token. The instruction string is
 *  whitespace-padded (` PAGE `) and may carry switches (`DATE \@ "M/d/yyyy"`),
 *  so we match on the leading keyword. */
export function fieldToken(instr: string): string {
	const upper = instr.trim().toUpperCase();
	if (upper.startsWith("PAGE")) return "{page}";
	if (upper.startsWith("NUMPAGES")) return "{pages}";
	if (upper.startsWith("DATE")) return "{date}";
	if (upper.startsWith("TIME")) return "{time}";
	if (upper.startsWith("STYLEREF")) {
		const match = instr.match(/STYLEREF\s+"([^"]+)"/i);
		return match ? `{styleref:${match[1]}}` : "{styleref}";
	}
	if (upper.startsWith("FILENAME")) return "{filename}";
	if (upper.startsWith("TITLE")) return "{title}";
	if (upper.startsWith("AUTHOR")) return "{author}";
	return "{field}";
}
