import type { DocView } from "../ast/doc-view";
import { w } from "../jsx";
import type { XmlNode } from "../parser";
import { ensureCustomStyle, ensureStyle } from "../styles";

/** Provision the styles a code block needs in `styles.xml`. Always defines
 *  `Code` (character) and `CodeBlock` (paragraph) from the baseline catalog.
 *  When `language` is given, additionally defines `CodeBlock-LANG` (basedOn
 *  `CodeBlock`) so the language survives round-trip: the reader recovers it
 *  from the pStyle suffix and renders the fenced block with the right tag.
 *  Word/LibreOffice see a normal paragraph style that inherits from CodeBlock
 *  — same rendering, plus a Styles-pane entry that lists the language. */
export function ensureCodeBlockStyles(
	view: DocView,
	language: string | undefined,
): void {
	ensureStyle(view, "Code");
	ensureStyle(view, "CodeBlock");
	if (!language) return;
	const styleId = codeBlockStyleIdFor(language);
	ensureCustomStyle(view, styleId, () => buildLanguageStyle(styleId, language));
}

function buildLanguageStyle(styleId: string, language: string): XmlNode {
	return (
		<w.style w-type="paragraph" w-styleId={styleId}>
			<w.name w-val={`Code Block (${language})`} />
			<w.basedOn w-val="CodeBlock" />
			<w.next w-val="Normal" />
			<w.qFormat />
		</w.style>
	);
}

/** pStyle id for a code block in the given language. Returns the bare
 *  `CodeBlock` for unlanguaged blocks — matches the baseline style and
 *  keeps the styles.xml footprint minimal. */
export function codeBlockStyleIdFor(language: string | undefined): string {
	return language ? `CodeBlock-${language}` : "CodeBlock";
}

/** Inverse of `codeBlockStyleIdFor`. Returns the language suffix for a
 *  `CodeBlock-LANG` style id; `undefined` for the bare `CodeBlock` or any
 *  non-code-block style. Used by `read --markdown` to emit a fenced block
 *  with the right language tag. */
export function codeBlockLanguageFromStyleId(
	styleId: string | undefined,
): string | undefined {
	if (!styleId) return undefined;
	if (!styleId.startsWith("CodeBlock-")) return undefined;
	const suffix = styleId.slice("CodeBlock-".length);
	return suffix.length > 0 ? suffix : undefined;
}

/** True when `styleId` is `CodeBlock` or any `CodeBlock-LANG` variant. The
 *  markdown render groups consecutive paragraphs with this style into one
 *  fenced block. */
export function isCodeBlockStyleId(styleId: string | undefined): boolean {
	if (!styleId) return false;
	return styleId === "CodeBlock" || styleId.startsWith("CodeBlock-");
}
