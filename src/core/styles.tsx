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
	| "FootnoteText"
	| "EndnoteReference"
	| "EndnoteText";

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
 * Used by `insert --style` / `edit --style` (via `isBaselineStyle`) so a
 * referenced baseline style is also defined, and by the S8 markdown walker. */
export function ensureStyle(view: DocView, styleId: BaselineStyleId): void {
	const tree = ensureStylesPart(view);
	const root = XmlNode.findRoot(tree, "w:styles");
	if (!root) throw new Error("expected <w:styles> root in stylesTree");
	if (styleId !== "Normal") ensureStyleNode(root, "Normal");
	ensureStyleNode(root, styleId);
}

/** Whether `styleId` is one of the baseline styles `ensureStyle` can define.
 * Lets callers provision a `--style` reference only when we have a definition
 * for it (custom/unknown styles are referenced but left for the doc to define). */
export function isBaselineStyle(styleId: string): styleId is BaselineStyleId {
	return Object.hasOwn(BASELINE, styleId);
}

/** Define a referenced paragraph/character style if it's a baseline we know
 * how to provision; no-op for undefined or custom styles. The convenience
 * `insert --style` / `edit --style` call so a `--style Heading2` reference
 * also lands a Heading2 definition in styles.xml (otherwise Word renders it
 * as Normal). */
export function ensureReferencedStyle(
	view: DocView,
	styleId: string | undefined,
): void {
	if (styleId && isBaselineStyle(styleId)) ensureStyle(view, styleId);
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

/** The baseline style catalog: id → a builder that emits the `<w:style>`
 * definition. Each builder renders one of the style components below. */
const BASELINE: Record<BaselineStyleId, () => XmlNode> = {
	Normal: () => <NormalStyle />,
	Heading1: () => (
		<HeadingStyle
			styleId="Heading1"
			displayName="heading 1"
			outlineLevel={0}
			sizeHalfPoints={32}
		/>
	),
	Heading2: () => (
		<HeadingStyle
			styleId="Heading2"
			displayName="heading 2"
			outlineLevel={1}
			sizeHalfPoints={28}
		/>
	),
	Heading3: () => (
		<HeadingStyle
			styleId="Heading3"
			displayName="heading 3"
			outlineLevel={2}
			sizeHalfPoints={26}
		/>
	),
	Heading4: () => (
		<HeadingStyle
			styleId="Heading4"
			displayName="heading 4"
			outlineLevel={3}
			sizeHalfPoints={24}
		/>
	),
	Heading5: () => (
		<HeadingStyle
			styleId="Heading5"
			displayName="heading 5"
			outlineLevel={4}
			sizeHalfPoints={22}
		/>
	),
	Heading6: () => (
		<HeadingStyle
			styleId="Heading6"
			displayName="heading 6"
			outlineLevel={5}
			sizeHalfPoints={20}
		/>
	),
	Quote: () => <QuoteStyle />,
	IntenseQuote: () => <IntenseQuoteStyle />,
	Code: () => <CodeStyle />,
	CodeBlock: () => <CodeBlockStyle />,
	ListParagraph: () => <ListParagraphStyle />,
	FootnoteReference: () => <FootnoteReferenceStyle />,
	FootnoteText: () => <FootnoteTextStyle />,
	EndnoteReference: () => <EndnoteReferenceStyle />,
	EndnoteText: () => <EndnoteTextStyle />,
};

function NormalStyle(): XmlNode {
	return (
		<w.style w-type="paragraph" w-default="1" w-styleId="Normal">
			<w.name w-val="Normal" />
			<w.qFormat />
		</w.style>
	);
}

function HeadingStyle({
	styleId,
	displayName,
	outlineLevel,
	sizeHalfPoints,
}: {
	styleId: BaselineStyleId;
	displayName: string;
	outlineLevel: number;
	sizeHalfPoints: number;
}): XmlNode {
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

function QuoteStyle(): XmlNode {
	return (
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
	);
}

function IntenseQuoteStyle(): XmlNode {
	return (
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
	);
}

function CodeStyle(): XmlNode {
	return (
		<w.style w-type="character" w-styleId="Code">
			<w.name w-val="Code" />
			<w.qFormat />
			<w.rPr>
				<w.rFonts w-ascii="Courier New" w-hAnsi="Courier New" />
			</w.rPr>
		</w.style>
	);
}

function CodeBlockStyle(): XmlNode {
	return (
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
	);
}

function ListParagraphStyle(): XmlNode {
	return (
		<w.style w-type="paragraph" w-styleId="ListParagraph">
			<w.name w-val="List Paragraph" />
			<w.basedOn w-val="Normal" />
			<w.qFormat />
			<w.pPr>
				<w.ind w-left="720" />
				{/* Collapse Normal's spacing.after between adjacent items of this
				    style — matches Word's own List Paragraph definition. */}
				<w.contextualSpacing />
			</w.pPr>
		</w.style>
	);
}

function FootnoteReferenceStyle(): XmlNode {
	return (
		<w.style w-type="character" w-styleId="FootnoteReference">
			<w.name w-val="footnote reference" />
			<w.rPr>
				<w.vertAlign w-val="superscript" />
			</w.rPr>
		</w.style>
	);
}

function FootnoteTextStyle(): XmlNode {
	return (
		<w.style w-type="paragraph" w-styleId="FootnoteText">
			<w.name w-val="footnote text" />
			<w.basedOn w-val="Normal" />
			<w.rPr>
				<w.sz w-val="20" />
			</w.rPr>
		</w.style>
	);
}

function EndnoteReferenceStyle(): XmlNode {
	return (
		<w.style w-type="character" w-styleId="EndnoteReference">
			<w.name w-val="endnote reference" />
			<w.rPr>
				<w.vertAlign w-val="superscript" />
			</w.rPr>
		</w.style>
	);
}

function EndnoteTextStyle(): XmlNode {
	return (
		<w.style w-type="paragraph" w-styleId="EndnoteText">
			<w.name w-val="endnote text" />
			<w.basedOn w-val="Normal" />
			<w.rPr>
				<w.sz w-val="20" />
			</w.rPr>
		</w.style>
	);
}
