import type { Root, RootContent } from "mdast";

/** Re-parse the content after a LEADING `docx:` hint comment.
 *
 * CommonMark opens an HTML block at any line that begins with `<!--` and runs
 * it to the end of the line holding `-->`, so a line `read` prints with an
 * inline zero-width hint first — `<!-- docx:chart -->Figure 2: sales <!-- p1 -->`
 * — parsed as ONE html block. The importer drops html blocks (every comment is
 * a hint, never content), which silently took the text after the hint with it:
 * a copied line handed back to `edit` replaced the paragraph with nothing.
 *
 * So an html block that STARTS with one or more `docx:` comments and still
 * holds text after them is replaced by the parse of that remainder — the hint
 * goes, the content stays, exactly as if the hint were mid-line (where it is
 * already an inline html node the walker drops). Only `docx:` comments qualify:
 * an author's own `<!-- … -->` keeps CommonMark's meaning. A line that is ONLY
 * comments (`<!-- docx:p p1 page-break -->`) is left as the dropped html block. */
export function reparseHintLedLines(
	tree: Root,
	parse: (source: string) => Root,
): void {
	reparseContainer(tree, parse);
}

function reparseContainer(
	node: { children?: unknown[] },
	parse: (source: string) => Root,
): void {
	if (!Array.isArray(node.children)) return;
	const children = node.children as RootContent[];
	const out: RootContent[] = [];
	for (const child of children) {
		const remainder =
			child.type === "html" ? hintLedRemainder(child.value) : null;
		if (remainder === null) {
			if (BLOCK_CONTAINERS.has(child.type)) {
				reparseContainer(child as { children?: unknown[] }, parse);
			}
			out.push(child);
			continue;
		}
		out.push(...parse(remainder).children);
	}
	node.children = out;
}

/** Block-level nodes whose children are blocks (where an html BLOCK can sit). */
const BLOCK_CONTAINERS: ReadonlySet<string> = new Set([
	"blockquote",
	"listItem",
	"footnoteDefinition",
	"list",
]);

const LEADING_HINTS = /^\s*(?:<!--\s*docx:[\s\S]*?-->[ \t]*)+/;
const ANY_COMMENT = /<!--[\s\S]*?-->/g;

/** The source after the leading `docx:` comments, or null when the block
 *  doesn't start with one or holds nothing but comments after it. */
function hintLedRemainder(value: string): string | null {
	const match = value.match(LEADING_HINTS);
	if (!match) return null;
	const remainder = value.slice(match[0].length);
	if (remainder.replace(ANY_COMMENT, "").trim() === "") return null;
	return remainder;
}
