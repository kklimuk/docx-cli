import { w } from "../../jsx";
import { XmlNode } from "../../parser";
import type { ContentTypesView } from "./content-types";
import type { Pkg } from "./package";
import type { RelationshipsView } from "./relationships";

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
	| "QuoteListParagraph"
	| "Hyperlink"
	| "FootnoteReference"
	| "FootnoteText"
	| "EndnoteReference"
	| "EndnoteText";

export function isBaselineStyle(styleId: string): styleId is BaselineStyleId {
	return Object.hasOwn(BASELINE, styleId);
}

const STYLES_PART_NAME = "word/styles.xml";
const STYLES_RELATIONSHIP_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const STYLES_CONTENT_TYPE =
	"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";

const EMPTY_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;

export class StylesView {
	tree: XmlNode[];

	constructor(tree: XmlNode[] = XmlNode.parse(EMPTY_STYLES_XML)) {
		this.tree = tree;
	}

	/** Load this view from a package; returns undefined if the part is absent. */
	static async fromPackage(pkg: Pkg): Promise<StylesView | undefined> {
		const tree = await pkg.readPart(STYLES_PART_NAME);
		return tree ? new StylesView(tree) : undefined;
	}

	/** Parse a view from raw XML; returns undefined if the input is absent. */
	static fromXml(xml: string | undefined): StylesView | undefined {
		return xml ? new StylesView(XmlNode.parse(xml)) : undefined;
	}

	/** Serialize this view's tree into the package's `word/styles.xml`. */
	writeTo(pkg: Pkg): void {
		pkg.writeText(STYLES_PART_NAME, XmlNode.serialize(this.tree));
	}

	/** Mint the styles relationship + content-type override on the containing
	 * package and return a fresh empty view. Idempotent on the relationship
	 * target. Called by `Document.ensureStyles()`. */
	static register(deps: {
		relationships: RelationshipsView;
		contentTypes: ContentTypesView;
	}): StylesView {
		if (!deps.relationships.hasTarget("styles.xml")) {
			deps.relationships.add(STYLES_RELATIONSHIP_TYPE, "styles.xml");
		}
		deps.contentTypes.registerPart(STYLES_PART_NAME, STYLES_CONTENT_TYPE);
		return new StylesView();
	}

	listStyleIds(): string[] {
		const root = XmlNode.findRoot(this.tree, "w:styles");
		if (!root) return [];
		const out: string[] = [];
		for (const child of root.findChildren("w:style")) {
			const id = child.getAttribute("w:styleId");
			if (id) out.push(id);
		}
		return out;
	}

	hasStyle(styleId: string): boolean {
		return this.getStyle(styleId) !== undefined;
	}

	getStyle(styleId: string): XmlNode | undefined {
		const root = XmlNode.findRoot(this.tree, "w:styles");
		if (!root) return undefined;
		return root
			.findChildren("w:style")
			.find((child) => child.getAttribute("w:styleId") === styleId);
	}

	/** Ensure the named baseline style is defined. Seeds `Normal` first so any
	 * `basedOn` references resolve. Subsequent calls for the same id are no-ops.
	 *
	 * Used by `insert --style` / `edit --style` (via `isBaselineStyle`) so a
	 * referenced baseline style is also defined, and by the markdown walker. */
	ensureStyle(styleId: BaselineStyleId): void {
		const root = this.ensureStylesRoot();
		if (styleId !== "Normal") this.#ensureStyleNode(root, "Normal");
		this.#ensureStyleNode(root, styleId);
	}

	/** Define a referenced paragraph/character style if it's a baseline we know
	 * how to provision; no-op for undefined or custom styles. Convenience for
	 * `insert --style` / `edit --style` so a `--style Heading2` reference also
	 * lands a Heading2 definition (otherwise Word renders it as Normal). */
	ensureReferencedStyle(styleId: string | undefined): void {
		if (styleId && isBaselineStyle(styleId)) this.ensureStyle(styleId);
	}

	/** Provision the baseline character styles referenced by any run's `runStyle`.
	 * Mirrors `ensureReferencedStyle` for paragraphs but walks a list of runs.
	 * Accepts a heterogeneous `Run[]`-like iterable — only TextRun carries
	 * `runStyle`, but the duck-type read here keeps the caller from filtering. */
	ensureReferencedRunStyles(runs: Iterable<unknown>): void {
		const seen = new Set<string>();
		for (const run of runs) {
			const styleId =
				typeof run === "object" && run !== null && "runStyle" in run
					? (run as { runStyle?: unknown }).runStyle
					: undefined;
			if (typeof styleId !== "string" || seen.has(styleId)) continue;
			seen.add(styleId);
			this.ensureReferencedStyle(styleId);
		}
	}

	/** Provision a custom (non-baseline) `<w:style>` whose definition the caller
	 * supplies via `build`. No-op if a style with `styleId` already exists. Seeds
	 * Normal first so any basedOn references in `build()` resolve. Used by the
	 * code-block package to provision `CodeBlock-LANG` styles on demand — the
	 * language is metadata stored in the styleId itself. */
	ensureCustomStyle(styleId: string, build: () => XmlNode): void {
		const root = this.ensureStylesRoot();
		this.#ensureStyleNode(root, "Normal");
		if (this.hasStyle(styleId)) return;
		root.children.push(build());
	}

	private ensureStylesRoot(): XmlNode {
		const root = XmlNode.findRoot(this.tree, "w:styles");
		if (!root) throw new Error("expected <w:styles> root in styles tree");
		return root;
	}

	#ensureStyleNode(stylesRoot: XmlNode, styleId: BaselineStyleId): void {
		const exists = stylesRoot
			.findChildren("w:style")
			.some((child) => child.getAttribute("w:styleId") === styleId);
		if (exists) return;
		stylesRoot.children.push(BASELINE[styleId]());
	}
}

/** The baseline style catalog: id → a builder that emits the `<w:style>`
 * definition. Each builder renders one of the style components below. */
const BASELINE: Record<BaselineStyleId, () => XmlNode> = {
	Normal: () => <NormalStyle />,
	// Modern Word (Office 365 / 2013+) heading defaults. Size + bold/italic
	// + color = visual hierarchy. Color cues (`2E74B5` mid-blue, `1F4D78`
	// darker blue) carry most of the differentiation since H3–H6 are all
	// 11pt; the bold/italic mix discriminates the lower levels.
	Heading1: () => (
		<HeadingStyle
			styleId="Heading1"
			displayName="heading 1"
			outlineLevel={0}
			sizeHalfPoints={32}
			bold
			color="2E74B5"
		/>
	),
	Heading2: () => (
		<HeadingStyle
			styleId="Heading2"
			displayName="heading 2"
			outlineLevel={1}
			sizeHalfPoints={26}
			bold
			color="2E74B5"
		/>
	),
	Heading3: () => (
		<HeadingStyle
			styleId="Heading3"
			displayName="heading 3"
			outlineLevel={2}
			sizeHalfPoints={24}
			bold
			color="1F4D78"
		/>
	),
	Heading4: () => (
		<HeadingStyle
			styleId="Heading4"
			displayName="heading 4"
			outlineLevel={3}
			sizeHalfPoints={22}
			bold
			italic
			color="2E74B5"
		/>
	),
	Heading5: () => (
		<HeadingStyle
			styleId="Heading5"
			displayName="heading 5"
			outlineLevel={4}
			sizeHalfPoints={22}
			color="2E74B5"
		/>
	),
	Heading6: () => (
		<HeadingStyle
			styleId="Heading6"
			displayName="heading 6"
			outlineLevel={5}
			sizeHalfPoints={22}
			italic
			color="1F4D78"
		/>
	),
	Quote: () => <QuoteStyle />,
	IntenseQuote: () => <IntenseQuoteStyle />,
	Code: () => <CodeStyle />,
	CodeBlock: () => <CodeBlockStyle />,
	ListParagraph: () => <ListParagraphStyle />,
	QuoteListParagraph: () => <QuoteListParagraphStyle />,
	Hyperlink: () => <HyperlinkStyle />,
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
	bold,
	italic,
	color,
}: {
	styleId: BaselineStyleId;
	displayName: string;
	outlineLevel: number;
	sizeHalfPoints: number;
	bold?: boolean;
	italic?: boolean;
	/** Hex color without `#` (e.g. `2E74B5`). Optional — Word's H5 is the
	 *  uncolored case among the modern defaults. */
	color?: string;
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
				{/* `<w:rFonts>` per ECMA-376 §17.3.2.26 sits before bold/italic in
				   `<w:rPr>`. Word's modern headings use Calibri Light (one font
				   for all heading levels); apply via `w:asciiTheme="majorHAnsi"`
				   if/when we wire up theme inheritance, but a direct ascii/hAnsi
				   reference is portable across docs without a theme part. */}
				<w.rFonts w-ascii="Calibri Light" w-hAnsi="Calibri Light" />
				{bold && <w.b />}
				{italic && <w.i />}
				{color && <w.color w-val={color} />}
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

function QuoteListParagraphStyle(): XmlNode {
	// List items inside a markdown blockquote. Extends ListParagraph so the
	// numbering machinery (numId / lvlText / bullet glyphs) keeps working,
	// and adds italic to match the visual treatment of `Quote`. The actual
	// left indent comes from a paragraph-level `<w:ind>` the markdown walker
	// emits — that way nesting depth is encoded as `720 * depth` twips and
	// the read side recovers `paragraph.quoteDepth` from the indent value.
	return (
		<w.style w-type="paragraph" w-styleId="QuoteListParagraph">
			<w.name w-val="Quote List Paragraph" />
			<w.basedOn w-val="ListParagraph" />
			<w.qFormat />
			<w.rPr>
				<w.i />
			</w.rPr>
		</w.style>
	);
}

function HyperlinkStyle(): XmlNode {
	// Word's canonical Hyperlink character style: theme color 0563C1 (blue
	// hyperlink), single underline. The inline walker applies `rStyle` to
	// every text run inside `<w:hyperlink>` so anchors are visually obvious
	// without an explicit color/underline flag on each run.
	return (
		<w.style w-type="character" w-styleId="Hyperlink">
			<w.name w-val="Hyperlink" />
			<w.basedOn w-val="DefaultParagraphFont" />
			<w.uiPriority w-val="99" />
			<w.unhideWhenUsed />
			<w.rPr>
				<w.color w-val="0563C1" />
				<w.u w-val="single" />
			</w.rPr>
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
