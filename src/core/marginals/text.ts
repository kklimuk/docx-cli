import { XmlNode } from "../parser";

/** Extract a header/footer part's text for the AST `Marginal.text`, with
 *  `<w:fldSimple>` fields rendered as readable tokens (`{page}`, `{date}`,
 *  `{styleref:Heading 1}`, …) rather than their cached value. Paragraphs join
 *  with newlines; an inter-zone `<w:tab/>` becomes a literal tab. Pure read
 *  helper (no JSX), shared by the AST reader. */
export function marginalText(tree: XmlNode[]): string {
	const root =
		XmlNode.findRoot(tree, "w:hdr") ?? XmlNode.findRoot(tree, "w:ftr");
	if (!root) return "";
	const paragraphs: string[] = [];
	for (const child of root.children) {
		if (child.tag !== "w:p") continue;
		paragraphs.push(paragraphText(child));
	}
	return paragraphs.join("\n").trim();
}

function paragraphText(paragraph: XmlNode): string {
	let out = "";
	for (const child of paragraph.children) out += nodeText(child);
	return out;
}

function nodeText(node: XmlNode): string {
	if (node.tag === "w:fldSimple") {
		return fieldToken(node.getAttribute("w:instr") ?? "");
	}
	if (node.tag === "w:r") {
		let out = "";
		for (const child of node.children) {
			if (child.tag === "w:t") out += child.collectText();
			else if (child.tag === "w:tab") out += "\t";
		}
		return out;
	}
	// Run-bearing wrappers (a future header hyperlink) — descend.
	if (node.tag === "w:hyperlink") {
		let out = "";
		for (const child of node.children) out += nodeText(child);
		return out;
	}
	return "";
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
