import type { Paragraph, PhrasingContent, Root } from "mdast";
import type { Math as MdMath } from "mdast-util-math";

/** Turn every top-level paragraph that is ONLY `$$…$$` into a display `math`
 *  block. `read` prints a display equation as `$$latex$$` on one line (its
 *  locator on the next), but remark-math parses a one-line `$$…$$` as INLINE
 *  math — so an agent copying that line back into `edit`/`insert`/`create`
 *  demoted the block to an inline equation, and a multi-line
 *  `\begin{gathered}` block collapsed into one inline `<m:oMath>`. A `$$`
 *  alone in its paragraph is display math in every renderer that shows it
 *  (GitHub, Obsidian, KaTeX), so it is here too. A `$$…$$` that shares its
 *  paragraph with other text stays inline. */
export function promoteDisplayMath(tree: Root, source: string): void {
	tree.children = tree.children.map((block) =>
		block.type === "paragraph" ? (displayMath(block, source) ?? block) : block,
	);
}

function displayMath(paragraph: Paragraph, source: string): MdMath | undefined {
	// Whitespace and `<!-- pN -->` locator comments are dropped on import, so
	// they don't count as other content.
	const content = paragraph.children.filter((child) => !isIgnorable(child));
	const [only] = content;
	if (content.length !== 1 || only?.type !== "inlineMath") return undefined;
	const start = only.position?.start.offset;
	if (start === undefined || !source.startsWith("$$", start)) return undefined;
	return { type: "math", value: only.value, position: paragraph.position };
}

function isIgnorable(child: PhrasingContent): boolean {
	if (child.type === "text") return child.value.trim() === "";
	return child.type === "html" && child.value.trim().startsWith("<!--");
}
