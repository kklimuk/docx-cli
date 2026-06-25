import {
	applyParagraphPropsToPPr,
	insertRprChildInOrder,
	type ParagraphOptions,
} from "../../blocks";
import { applyRunFormatToRpr, type RunFormat } from "../../edit/set-formatting";
import { w } from "../../jsx";
import { XmlNode } from "../../parser";
import type { ContentTypesView } from "./content-types";
import type { Pkg } from "./package";
import type { RelationshipsView } from "./relationships";

export type BaselineStyleId =
	| "Normal"
	| "Title"
	| "Subtitle"
	| "Heading1"
	| "Heading2"
	| "Heading3"
	| "Heading4"
	| "Heading5"
	| "Heading6"
	| "Heading7"
	| "Heading8"
	| "Heading9"
	| "Quote"
	| "IntenseQuote"
	| "Code"
	| "CodeBlock"
	| "ListParagraph"
	| "QuoteListParagraph"
	| "Caption"
	| "Hyperlink"
	| "FootnoteReference"
	| "FootnoteText"
	| "EndnoteReference"
	| "EndnoteText";

export function isBaselineStyle(styleId: string): styleId is BaselineStyleId {
	return Object.hasOwn(BASELINE, styleId);
}

/** The built-in style catalog docx-cli can provision on demand — every id that
 *  `insert --style` / `edit --style` auto-defines (via `ensureStyle`) even when
 *  the document doesn't yet contain it. `styles --catalog` lists these so agents
 *  can discover valid `--style` values (Title, Subtitle, Heading1–9, Quote, …)
 *  before applying them, instead of guessing. Returns freshly-built `<w:style>`
 *  nodes; the caller reads id / type / name off each. */
export function baselineCatalog(): XmlNode[] {
	return (Object.keys(BASELINE) as BaselineStyleId[]).map((id) =>
		BASELINE[id](),
	);
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

	/** Default run font size in half-points, from `<w:docDefaults><w:rPrDefault>`
	 * (falling back to the Normal style's run properties). `read` uses it as the
	 * baseline size when no stronger per-run majority emerges, so a Word doc that
	 * stamps an explicit `<w:sz>` matching the default on most runs still reads
	 * clean (the note declares it once; the runs omit it). */
	defaultSizeHalfPoints(): number | undefined {
		const root = XmlNode.findRoot(this.tree, "w:styles");
		if (!root) return undefined;
		const docDefault = root
			.findChild("w:docDefaults")
			?.findChild("w:rPrDefault")
			?.findChild("w:rPr")
			?.findChild("w:sz")
			?.getAttribute("w:val");
		if (docDefault) return Number(docDefault);
		const normal = this.getStyle("Normal")
			?.findChild("w:rPr")
			?.findChild("w:sz")
			?.getAttribute("w:val");
		return normal ? Number(normal) : undefined;
	}

	/** Default run font (explicit `w:ascii`) from `<w:docDefaults><w:rPrDefault>`,
	 * falling back to the Normal style's run properties. The counterpart to
	 * `defaultSizeHalfPoints`. `read` surfaces it in the `docx:base` note when it
	 * DEVIATES from the canonical template default (i.e. `set-default-font` ran),
	 * so the document font is observable on the next read. Theme-only docDefaults
	 * (no explicit `w:ascii`) return undefined — we don't resolve the theme here;
	 * `set-default-font` always writes an explicit ascii, so its effect is caught. */
	defaultFont(): string | undefined {
		const root = XmlNode.findRoot(this.tree, "w:styles");
		if (!root) return undefined;
		const docDefault = root
			.findChild("w:docDefaults")
			?.findChild("w:rPrDefault")
			?.findChild("w:rPr")
			?.findChild("w:rFonts")
			?.getAttribute("w:ascii");
		if (docDefault) return docDefault;
		return this.getStyle("Normal")
			?.findChild("w:rPr")
			?.findChild("w:rFonts")
			?.getAttribute("w:ascii");
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

	/** Set the document default run font: `<w:docDefaults>/<w:rPrDefault>/<w:rPr>/
	 *  <w:rFonts>`, creating the docDefaults scaffold if absent. Writes explicit
	 *  `w:ascii`/`w:hAnsi`/`w:cs` (and drops any theme reference) so the font wins
	 *  for every run that doesn't pin its own — the lowest, broadest font layer. */
	setDefaultFont(fontName: string): void {
		const rPr = this.#ensureDocDefaultsRPr();
		let rFonts = rPr.findChild("w:rFonts");
		if (!rFonts) {
			rFonts = XmlNode.element("w:rFonts");
			insertRprChildInOrder(rPr, rFonts);
		}
		applyRunFont(rFonts, fontName);
	}

	/** Set the document default run size (`<w:sz>`/`<w:szCs>` in docDefaults), in
	 *  half-points. Inserted at the canonical CT_RPr slot (after color/kern/…, not
	 *  merely after `<w:rFonts>`) so a docDefaults rPr that already carries other
	 *  run properties stays Word-valid. */
	setDefaultSizeHalfPoints(halfPoints: number): void {
		const rPr = this.#ensureDocDefaultsRPr();
		const value = String(halfPoints);
		let sz = rPr.findChild("w:sz");
		if (!sz) {
			sz = XmlNode.element("w:sz");
			insertRprChildInOrder(rPr, sz);
		}
		sz.setAttribute("w:val", value);
		let szCs = rPr.findChild("w:szCs");
		if (!szCs) {
			szCs = XmlNode.element("w:szCs");
			insertRprChildInOrder(rPr, szCs);
		}
		szCs.setAttribute("w:val", value);
	}

	/** Style ids whose own definition pins an explicit (non-theme) `w:ascii` font.
	 *  These OVERRIDE the document default, so a default-font change leaves them
	 *  looking different — the CLI names them so the agent knows what `--all`
	 *  would additionally touch. Theme-referencing styles aren't listed: they
	 *  follow the theme, which `Fonts` repoints alongside the default. */
	explicitFontStyleIds(): string[] {
		const root = XmlNode.findRoot(this.tree, "w:styles");
		if (!root) return [];
		const out: string[] = [];
		for (const style of root.findChildren("w:style")) {
			const ascii = style
				.findChild("w:rPr")
				?.findChild("w:rFonts")
				?.getAttribute("w:ascii");
			if (!ascii) continue;
			const id = style.getAttribute("w:styleId");
			if (id) out.push(id);
		}
		return out;
	}

	/** Repoint every style definition's explicit `<w:rFonts>` to `fontName` (the
	 *  `--all` path). Returns the count changed. */
	overrideStyleFonts(fontName: string): number {
		const root = XmlNode.findRoot(this.tree, "w:styles");
		if (!root) return 0;
		let count = 0;
		for (const style of root.findChildren("w:style")) {
			const rFonts = style.findChild("w:rPr")?.findChild("w:rFonts");
			if (!rFonts) continue;
			applyRunFont(rFonts, fontName);
			count++;
		}
		return count;
	}

	/** Apply run + paragraph formatting onto an EXISTING `<w:style>` definition
	 *  (the `styles set` verb). The style-definition twin of `setFormatting`
	 *  (runs) / `applyParagraphOptionsInPlace` (paragraphs): it find-or-creates
	 *  the style's `<w:rPr>`/`<w:pPr>` at their CT_Style slots, then applies the
	 *  SAME rPr/pPr vocabulary the body edits use. Optional `name`/`basedOn`/`next`
	 *  update the metadata children. In-place — any property the spec doesn't set is
	 *  preserved (and a paragraph using the style with its OWN direct override keeps
	 *  winning; we never touch the body). Style edits are deliberately UNTRACKED —
	 *  Word itself mutates styles.xml directly even under Track Changes. No-op if
	 *  `styleId` doesn't exist (the caller validates existence first for a clean
	 *  error and owns the `w:type` lookup it needs for the ack). */
	setStyleFormatting(styleId: string, spec: StyleSpec): void {
		const style = this.getStyle(styleId);
		if (!style) return;
		applyStyleSpec(style, spec);
	}

	/** Provision a NEW custom `<w:style>` with the given metadata + formatting (the
	 *  `styles create` verb). No-op if `styleId` already exists (caller errors
	 *  first). Seeds Normal so a `basedOn` reference resolves. Paragraph styles
	 *  default `basedOn`/`next` to Normal and get a `<w:pPr>` for any paragraph
	 *  options; character styles take run formatting only (caller rejects pPr flags
	 *  on them). `<w:qFormat/>` is emitted so the style shows in Word's gallery. */
	createStyle(spec: {
		styleId: string;
		type: "paragraph" | "character";
		name?: string;
		basedOn?: string;
		next?: string;
		runFormat?: RunFormat;
		paragraphOptions?: ParagraphOptions;
	}): void {
		this.ensureCustomStyle(spec.styleId, () => {
			const style = XmlNode.element("w:style");
			style.setAttribute("w:type", spec.type);
			style.setAttribute("w:styleId", spec.styleId);
			insertStyleChildInOrder(style, XmlNode.element("w:qFormat"));
			applyStyleSpec(style, {
				name: spec.name ?? spec.styleId,
				basedOn:
					spec.basedOn ?? (spec.type === "paragraph" ? "Normal" : undefined),
				next: spec.next ?? (spec.type === "paragraph" ? "Normal" : undefined),
				runFormat: spec.runFormat,
				paragraphOptions: spec.paragraphOptions,
			});
			return style;
		});
	}

	/** Navigate to (creating if absent) `<w:docDefaults>/<w:rPrDefault>/<w:rPr>`.
	 *  `<w:docDefaults>` is the first child of `<w:styles>` per CT_Styles. */
	#ensureDocDefaultsRPr(): XmlNode {
		const root = this.ensureStylesRoot();
		let docDefaults = root.findChild("w:docDefaults");
		if (!docDefaults) {
			docDefaults = XmlNode.element("w:docDefaults");
			root.children.unshift(docDefaults);
		}
		let rPrDefault = docDefaults.findChild("w:rPrDefault");
		if (!rPrDefault) {
			rPrDefault = XmlNode.element("w:rPrDefault");
			docDefaults.children.unshift(rPrDefault); // before <w:pPrDefault>
		}
		let rPr = rPrDefault.findChild("w:rPr");
		if (!rPr) {
			rPr = XmlNode.element("w:rPr");
			rPrDefault.children.push(rPr);
		}
		return rPr;
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

/** Point a `<w:rFonts>` at `fontName` for the Latin/ASCII script: set
 *  `w:ascii`/`w:hAnsi`/`w:cs` and DROP any `w:asciiTheme`/`w:hAnsiTheme`/
 *  `w:cstheme` reference (an explicit font must beat the theme). East-Asian /
 *  complex-script fallbacks (`w:eastAsia`) are left alone. Shared by
 *  `setDefaultFont`, `overrideStyleFonts`, and the `Fonts` lens's body/note walk. */
export function applyRunFont(rFonts: XmlNode, fontName: string): void {
	rFonts.setAttribute("w:ascii", fontName);
	rFonts.setAttribute("w:hAnsi", fontName);
	rFonts.setAttribute("w:cs", fontName);
	delete rFonts.attributes["w:asciiTheme"];
	delete rFonts.attributes["w:hAnsiTheme"];
	delete rFonts.attributes["w:cstheme"];
}

/** The metadata + formatting a `styles set`/`create` applies onto a `<w:style>`.
 *  Each field is optional; only present fields are written. */
export type StyleSpec = {
	runFormat?: RunFormat;
	paragraphOptions?: ParagraphOptions;
	name?: string;
	basedOn?: string;
	next?: string;
};

/** Apply a `StyleSpec`'s metadata + formatting onto a `<w:style>` node, in place,
 *  honoring CT_Style child order. Shared by `setStyleFormatting` (an existing
 *  node) and `createStyle` (a fresh one). Run formatting reuses `applyRunFormatToRpr`
 *  and paragraph options reuse `applyParagraphPropsToPPr` — the exact appliers the
 *  body edits use, so a style's `--bold`/`--space-before`/… behave identically. */
function applyStyleSpec(style: XmlNode, spec: StyleSpec): void {
	if (spec.name !== undefined) putStyleMeta(style, "w:name", spec.name);
	if (spec.basedOn !== undefined)
		putStyleMeta(style, "w:basedOn", spec.basedOn);
	if (spec.next !== undefined) putStyleMeta(style, "w:next", spec.next);
	if (spec.paragraphOptions) {
		applyParagraphPropsToPPr(ensureStylePPr(style), spec.paragraphOptions);
	}
	if (spec.runFormat) {
		applyRunFormatToRpr(ensureStyleRPr(style), spec.runFormat);
	}
}

/** Replace-or-insert a single-`w:val` metadata child (`<w:name>`/`<w:basedOn>`/
 *  `<w:next>`) at its CT_Style slot. */
function putStyleMeta(style: XmlNode, tag: string, value: string): void {
	const existing = style.findChild(tag);
	if (existing) {
		existing.setAttribute("w:val", value);
		return;
	}
	const node = XmlNode.element(tag);
	node.setAttribute("w:val", value);
	insertStyleChildInOrder(style, node);
}

/** Find the style's `<w:pPr>`, or splice an empty one in at its CT_Style slot. */
function ensureStylePPr(style: XmlNode): XmlNode {
	const existing = style.findChild("w:pPr");
	if (existing) return existing;
	const created = XmlNode.element("w:pPr");
	insertStyleChildInOrder(style, created);
	return created;
}

/** Find the style's `<w:rPr>`, or splice an empty one in at its CT_Style slot. */
function ensureStyleRPr(style: XmlNode): XmlNode {
	const existing = style.findChild("w:rPr");
	if (existing) return existing;
	const created = XmlNode.element("w:rPr");
	insertStyleChildInOrder(style, created);
	return created;
}

/** CT_Style child order (ECMA-376 §17.7.4.17) — the subset we author. pPr precedes
 *  rPr, and both follow the style metadata. Splicing a fresh `<w:name>`/`<w:pPr>`/
 *  `<w:rPr>`/… at the right slot keeps Word from rejecting the style. */
const STYLE_CHILD_ORDER = [
	"w:name",
	"w:aliases",
	"w:basedOn",
	"w:next",
	"w:link",
	"w:autoRedefine",
	"w:hidden",
	"w:uiPriority",
	"w:semiHidden",
	"w:unhideWhenUsed",
	"w:qFormat",
	"w:locked",
	"w:personal",
	"w:personalCompose",
	"w:personalReply",
	"w:rsid",
	"w:pPr",
	"w:rPr",
	"w:tblPr",
	"w:trPr",
	"w:tcPr",
	"w:tblStylePr",
] as const;

/** Rank a style child by CT_Style position. Unknown tags rank just before pPr/rPr
 *  so any metadata we don't model stays ahead of the property blocks (the only
 *  ordering that affects validity). */
function styleChildRank(tag: string): number {
	const index = STYLE_CHILD_ORDER.indexOf(
		tag as (typeof STYLE_CHILD_ORDER)[number],
	);
	if (index >= 0) return index;
	return STYLE_CHILD_ORDER.indexOf("w:pPr") - 0.5;
}

/** Splice `child` into `style.children` at its canonical CT_Style position: before
 *  the first existing child that ranks after it. The CT_Style analog of
 *  `insertPprChildInOrder`/`insertRprChildInOrder`. */
function insertStyleChildInOrder(style: XmlNode, child: XmlNode): void {
	const rank = styleChildRank(child.tag);
	// Rank only against real element siblings: a pretty-printed styles.xml (as
	// Word/LibreOffice/third-party tools emit) keeps inter-element whitespace as
	// `#text` nodes, which rank as "unknown" and would otherwise shove the new
	// child to the front and break CT_Style order.
	const at = style.children.findIndex(
		(existing) => !existing.isText && styleChildRank(existing.tag) > rank,
	);
	if (at < 0) style.children.push(child);
	else style.children.splice(at, 0, child);
}

/** The baseline style catalog: id → a builder that emits the `<w:style>`
 * definition. Each builder renders one of the style components below. */
const BASELINE: Record<BaselineStyleId, () => XmlNode> = {
	Normal: () => <NormalStyle />,
	Title: () => <TitleStyle />,
	Subtitle: () => <SubtitleStyle />,
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
	// H7–H9 keep the 11pt body size (like H5/H6) and lean on color + italic for
	// the remaining hierarchy — matching Word's modern defaults where the deep
	// levels are visually subtle.
	Heading7: () => (
		<HeadingStyle
			styleId="Heading7"
			displayName="heading 7"
			outlineLevel={6}
			sizeHalfPoints={22}
			color="2E74B5"
		/>
	),
	Heading8: () => (
		<HeadingStyle
			styleId="Heading8"
			displayName="heading 8"
			outlineLevel={7}
			sizeHalfPoints={22}
			italic
			color="272727"
		/>
	),
	Heading9: () => (
		<HeadingStyle
			styleId="Heading9"
			displayName="heading 9"
			outlineLevel={8}
			sizeHalfPoints={22}
			color="272727"
		/>
	),
	Quote: () => <QuoteStyle />,
	IntenseQuote: () => <IntenseQuoteStyle />,
	Code: () => <CodeStyle />,
	CodeBlock: () => <CodeBlockStyle />,
	ListParagraph: () => <ListParagraphStyle />,
	QuoteListParagraph: () => <QuoteListParagraphStyle />,
	Caption: () => <CaptionStyle />,
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

function TitleStyle(): XmlNode {
	// Word's built-in Title: large (28pt) Calibri Light, dark blue. The
	// document's top-line heading — what an agent reaches for to open a doc.
	return (
		<w.style w-type="paragraph" w-styleId="Title">
			<w.name w-val="Title" />
			<w.basedOn w-val="Normal" />
			<w.next w-val="Normal" />
			<w.qFormat />
			<w.pPr>
				<w.spacing w-after="60" />
			</w.pPr>
			<w.rPr>
				<w.rFonts w-ascii="Calibri Light" w-hAnsi="Calibri Light" />
				<w.color w-val="1F3864" />
				<w.sz w-val="56" />
			</w.rPr>
		</w.style>
	);
}

function SubtitleStyle(): XmlNode {
	// Word's built-in Subtitle: 14pt grey Calibri Light, sits under a Title.
	return (
		<w.style w-type="paragraph" w-styleId="Subtitle">
			<w.name w-val="Subtitle" />
			<w.basedOn w-val="Normal" />
			<w.next w-val="Normal" />
			<w.qFormat />
			<w.pPr>
				<w.spacing w-after="160" />
			</w.pPr>
			<w.rPr>
				<w.rFonts w-ascii="Calibri Light" w-hAnsi="Calibri Light" />
				<w.color w-val="5A5A5A" />
				<w.sz w-val="28" />
			</w.rPr>
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

function CaptionStyle(): XmlNode {
	// Word's built-in "caption" paragraph style — the one Insert Caption applies
	// under figures/tables. Italic, smaller (9pt), accent-blue text, with a touch
	// of spacing. A caption-styled paragraph is what makes a figure label show up
	// in a Table of Figures, so emitting this (rather than ad-hoc italic text) is
	// what makes `--caption` produce a native, reference-able caption.
	return (
		<w.style w-type="paragraph" w-styleId="Caption">
			<w.name w-val="caption" />
			<w.basedOn w-val="Normal" />
			<w.next w-val="Normal" />
			<w.qFormat />
			<w.pPr>
				<w.spacing w-before="0" w-after="200" />
			</w.pPr>
			<w.rPr>
				<w.i />
				<w.color w-val="44546A" w-themeColor="text2" />
				<w.sz w-val="18" />
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
