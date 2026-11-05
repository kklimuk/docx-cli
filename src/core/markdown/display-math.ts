import type { Paragraph, PhrasingContent, Root } from "mdast";
import type { InlineMath, Math as MdMath } from "mdast-util-math";

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
	tree.children = tree.children.map((block) => {
		if (block.type !== "paragraph") return block;
		const promoted = displayMath(block, source);
		if (promoted) return promoted;
		markTrackedDisplayMath(block, source);
		return block;
	});
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

/** A tracked display equation reads under `--current` as
 *  `{--$$old$$--}{++$$new$$++}` — a paragraph of CriticMarkup, each wrapping
 *  ONE `$$…$$`. It can't become a `math` block (a block can't sit inside a
 *  `<w:ins>`), so each such inline math is flagged `data.display` instead and
 *  the inline walker emits it as `<m:oMathPara>` inside its `<w:del>`/`<w:ins>`
 *  — the shape a tracked display-equation replace writes. Without it, pasting
 *  that line back (or `create --from` a `--current` read) demoted the block to
 *  an inline equation. */
function markTrackedDisplayMath(paragraph: Paragraph, source: string): void {
	const content = paragraph.children.filter((child) => !isIgnorable(child));
	if (content.length === 0) return;
	const equations: InlineMath[] = [];
	for (const child of content) {
		if (child.type !== "criticInsert" && child.type !== "criticDelete") return;
		const inner = child.children.filter((node) => !isIgnorable(node));
		const [only] = inner;
		if (inner.length !== 1 || only?.type !== "inlineMath") return;
		const start = only.position?.start.offset;
		if (start === undefined || !source.startsWith("$$", start)) return;
		equations.push(only);
	}
	for (const equation of equations) {
		equation.data = { ...equation.data, display: true };
	}
}

function isIgnorable(child: PhrasingContent): boolean {
	if (child.type === "text") return child.value.trim() === "";
	return child.type === "html" && child.value.trim().startsWith("<!--");
}

declare module "mdast-util-math" {
	interface InlineMathData {
		/** Emit as a display block (`<m:oMathPara>`) — see `markTrackedDisplayMath`. */
		display?: boolean;
	}
}
