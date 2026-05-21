import type { DocView } from "./ast/doc-view";
import { w } from "./jsx";
import { registerPart } from "./package";
import { XmlNode } from "./parser";

export type BaselineStyleId =
	| "Normal"
	| "Heading1"
	| "Heading2"
	| "Heading3"
	| "Heading4"
	| "Heading5"
	| "Heading6"
	| "Quote"
	| "IntenseQuote"
	| "Code"
	| "CodeBlock"
	| "ListParagraph"
	| "FootnoteReference"
	| "FootnoteText";

const STYLES_PART = {
	partName: "word/styles.xml",
	contentType:
		"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
	relationshipType:
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
	target: "styles.xml",
};

const EMPTY_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;

/** Ensure the named paragraph/character style is defined in word/styles.xml.
 *
 * Creates the styles part itself if the package lacks one (and registers the
 * relationship + content-type override). For styles that derive from Normal,
 * Normal is also seeded so basedOn references resolve. Subsequent calls for
 * the same id are no-ops.
 *
 * Exposed for upcoming features (S6 footnotes, S8 markdown walker) — not
 * currently wired into any CLI verb. */
export function ensureStyle(view: DocView, styleId: BaselineStyleId): void {
	const tree = ensureStylesPart(view);
	const root = XmlNode.findRoot(tree, "w:styles");
	if (!root) throw new Error("expected <w:styles> root in stylesTree");
	if (styleId !== "Normal") ensureStyleNode(root, "Normal");
	ensureStyleNode(root, styleId);
}

function ensureStylesPart(view: DocView): XmlNode[] {
	if (view.stylesTree) return view.stylesTree;
	view.stylesTree = XmlNode.parse(EMPTY_STYLES_XML);
	registerPart(view.relationshipsTree, view.contentTypesTree, STYLES_PART);
	return view.stylesTree;
}

function ensureStyleNode(stylesRoot: XmlNode, styleId: BaselineStyleId): void {
	const exists = stylesRoot
		.findChildren("w:style")
		.some((child) => child.getAttribute("w:styleId") === styleId);
	if (exists) return;
	stylesRoot.children.push(BASELINE[styleId]());
}

const BASELINE: Record<BaselineStyleId, () => XmlNode> = {
	Normal: () => (
		<w.style w-type="paragraph" w-default="1" w-styleId="Normal">
			<w.name w-val="Normal" />
			<w.qFormat />
		</w.style>
	),
	Heading1: () => heading("Heading1", "heading 1", 0, 32),
	Heading2: () => heading("Heading2", "heading 2", 1, 28),
	Heading3: () => heading("Heading3", "heading 3", 2, 26),
	Heading4: () => heading("Heading4", "heading 4", 3, 24),
	Heading5: () => heading("Heading5", "heading 5", 4, 22),
	Heading6: () => heading("Heading6", "heading 6", 5, 20),
	Quote: () => (
		<w.style w-type="paragraph" w-styleId="Quote">
			<w.name w-val="Quote" />
			<w.basedOn w-val="Normal" />
			<w.next w-val="Normal" />
			<w.qFormat />
			<w.pPr>
				<w.spacing w-before="120" w-after="120" />
				<w.ind w-left="720" w-right="720" />
			</w.pPr>
			<w.rPr>
				<w.i />
			</w.rPr>
		</w.style>
	),
	IntenseQuote: () => (
		<w.style w-type="paragraph" w-styleId="IntenseQuote">
			<w.name w-val="Intense Quote" />
			<w.basedOn w-val="Normal" />
			<w.next w-val="Normal" />
			<w.qFormat />
			<w.pPr>
				{/* Per ECMA-376 §17.3.1.26 (CT_PPrBase), pBdr precedes spacing/ind. */}
				<w.pBdr>
					<w.bottom w-val="single" w-sz="4" w-space="4" w-color="auto" />
				</w.pBdr>
				<w.spacing w-before="120" w-after="120" />
				<w.ind w-left="720" w-right="720" />
			</w.pPr>
			<w.rPr>
				<w.b />
				<w.i />
			</w.rPr>
		</w.style>
	),
	Code: () => (
		<w.style w-type="character" w-styleId="Code">
			<w.name w-val="Code" />
			<w.qFormat />
			<w.rPr>
				<w.rFonts w-ascii="Courier New" w-hAnsi="Courier New" />
			</w.rPr>
		</w.style>
	),
	CodeBlock: () => (
		<w.style w-type="paragraph" w-styleId="CodeBlock">
			<w.name w-val="Code Block" />
			<w.basedOn w-val="Normal" />
			<w.next w-val="Normal" />
			<w.qFormat />
			<w.pPr>
				<w.spacing w-before="120" w-after="120" />
				<w.ind w-left="360" />
			</w.pPr>
			<w.rPr>
				<w.rFonts w-ascii="Courier New" w-hAnsi="Courier New" />
				<w.sz w-val="20" />
			</w.rPr>
		</w.style>
	),
	ListParagraph: () => (
		<w.style w-type="paragraph" w-styleId="ListParagraph">
			<w.name w-val="List Paragraph" />
			<w.basedOn w-val="Normal" />
			<w.qFormat />
			<w.pPr>
				<w.ind w-left="720" />
			</w.pPr>
		</w.style>
	),
	FootnoteReference: () => (
		<w.style w-type="character" w-styleId="FootnoteReference">
			<w.name w-val="footnote reference" />
			<w.rPr>
				<w.vertAlign w-val="superscript" />
			</w.rPr>
		</w.style>
	),
	FootnoteText: () => (
		<w.style w-type="paragraph" w-styleId="FootnoteText">
			<w.name w-val="footnote text" />
			<w.basedOn w-val="Normal" />
			<w.rPr>
				<w.sz w-val="20" />
			</w.rPr>
		</w.style>
	),
};

function heading(
	styleId: BaselineStyleId,
	displayName: string,
	outlineLevel: number,
	sizeHalfPoints: number,
): XmlNode {
	return (
		<w.style w-type="paragraph" w-styleId={styleId}>
			<w.name w-val={displayName} />
			<w.basedOn w-val="Normal" />
			<w.next w-val="Normal" />
			<w.qFormat />
			<w.pPr>
				<w.keepNext />
				<w.keepLines />
				<w.spacing w-before="240" w-after="60" />
				<w.outlineLvl w-val={String(outlineLevel)} />
			</w.pPr>
			<w.rPr>
				<w.b />
				<w.sz w-val={String(sizeHalfPoints)} />
			</w.rPr>
		</w.style>
	);
}
