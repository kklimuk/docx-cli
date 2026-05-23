import type { Nodes } from "hast";
import { common, createLowlight } from "lowlight";

/** A flat token list: each token is some text plus the color it should render
 *  in. Used by `paragraphs.tsx` to build per-line runs. */
export type CodeToken = { text: string; color?: string };

const lowlight = createLowlight(common);

/** Tokenize `code` for `language`. Returns `null` if the language isn't in
 *  the bundled `common` grammar set — caller falls back to the uncolored
 *  single-token path. Each leaf text node in the resulting hast tree becomes
 *  a token, attributed with the first color-mapped class found on any of its
 *  ancestor spans. Unknown highlight.js classes fall through with no color. */
export function highlightCode(
	language: string,
	code: string,
): CodeToken[] | null {
	if (!lowlight.registered(language)) return null;
	const tree = lowlight.highlight(language, code);
	const out: CodeToken[] = [];
	flatten(tree, undefined, out);
	return out;
}

function flatten(
	node: Nodes,
	inheritedColor: string | undefined,
	out: CodeToken[],
): void {
	if (node.type === "text") {
		const token: CodeToken = { text: node.value };
		if (inheritedColor) token.color = inheritedColor;
		out.push(token);
		return;
	}
	if (node.type === "root") {
		for (const child of node.children) flatten(child, inheritedColor, out);
		return;
	}
	if (node.type === "element") {
		// Pick the most specific color from this span's classes, falling back to
		// the inherited color if none match. We take the FIRST hit so a nested
		// span tagged `hljs-title function_` picks `hljs-title` (purple) over
		// the unmapped `function_`.
		const className = node.properties.className;
		const classes = Array.isArray(className) ? (className as string[]) : [];
		let color = inheritedColor;
		for (const cls of classes) {
			const mapped = COLOR_BY_CLASS[cls];
			if (mapped) {
				color = mapped;
				break;
			}
		}
		for (const child of node.children) flatten(child, color, out);
	}
	// Other node types (comment, doctype) — lowlight doesn't emit them, but
	// the `Nodes` union covers them; they fall through with no output.
}

/** GitHub-light inspired palette mapping highlight.js classes to hex colors
 *  (no `#`, OOXML stores 6-digit upper-case). Covers the major token kinds
 *  highlight.js emits across the bundled `common` grammars; unmapped classes
 *  fall through with no color (rendered in the default body color). The
 *  palette is intentionally narrow — distinguishing keywords from strings
 *  from comments is what makes code legible; finer gradations (e.g. distinct
 *  colors for `hljs-built_in` vs `hljs-type`) buy diminishing returns. */
const COLOR_BY_CLASS: Record<string, string> = {
	"hljs-keyword": "CF222E",
	"hljs-built_in": "0550AE",
	"hljs-type": "953800",
	"hljs-literal": "0550AE",
	"hljs-number": "0550AE",
	"hljs-string": "0A3069",
	"hljs-regexp": "0A3069",
	"hljs-comment": "6E7781",
	"hljs-doctag": "6E7781",
	"hljs-meta": "0A3069",
	"hljs-tag": "116329",
	"hljs-name": "0550AE",
	"hljs-attr": "0550AE",
	"hljs-attribute": "0550AE",
	"hljs-symbol": "0550AE",
	"hljs-title": "8250DF",
	"hljs-class": "953800",
	"hljs-function": "8250DF",
	"hljs-variable": "953800",
	"hljs-property": "953800",
	"hljs-section": "0550AE",
	"hljs-deletion": "82071E",
	"hljs-addition": "116329",
};
