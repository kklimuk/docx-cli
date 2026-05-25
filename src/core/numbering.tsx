import type { DocView } from "./ast/doc-view";
import { w } from "./jsx";
import { registerPart } from "./package";
import { XmlNode } from "./parser";

export type AbstractNumKind = "bullet" | "ordered";

const NUMBERING_PART = {
	partName: "word/numbering.xml",
	contentType:
		"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
	relationshipType:
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
	target: "numbering.xml",
};

const EMPTY_NUMBERING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;

/** Allocate a fresh `numId` for a new list of the given kind, pointing at the
 * package's shared bullet or ordered `abstractNum`. Each call returns a
 * distinct `numId` so multiple lists in the same document restart their
 * numbering independently (per OOXML convention: abstractNum is the style;
 * num is the per-list instance).
 *
 * Creates `word/numbering.xml` if absent (registering the relationship and
 * content-type override), and seeds the bullet/ordered `abstractNum` on
 * first use. Style provisioning for the `ListParagraph` pStyle is the
 * caller's responsibility — typically via `ensureStyle(view, "ListParagraph")`.
 *
 * Exposed for upcoming features (S8 markdown walker) — not currently wired
 * into any CLI verb. */
export function allocateNum(view: DocView, kind: AbstractNumKind): number {
	const tree = ensureNumberingPart(view);
	const root = XmlNode.findRoot(tree, "w:numbering");
	if (!root) throw new Error("expected <w:numbering> root in numberingTree");
	const abstractNumId = ensureAbstractNum(root, kind);
	const numId = nextNumId(root);
	root.children.push(
		<w.num w-numId={String(numId)}>
			<w.abstractNumId w-val={String(abstractNumId)} />
		</w.num>,
	);
	return numId;
}

function ensureNumberingPart(view: DocView): XmlNode[] {
	if (view.numberingTree) return view.numberingTree;
	view.numberingTree = XmlNode.parse(EMPTY_NUMBERING_XML);
	registerPart(view.relationshipsTree, view.contentTypesTree, NUMBERING_PART);
	return view.numberingTree;
}

function ensureAbstractNum(root: XmlNode, kind: AbstractNumKind): number {
	// Reuse keys on level-0 numFmt alone — sufficient for our own abstractNums,
	// but won't deduplicate identically-shaped abstractNums seeded by other
	// tools that happen to share a level-0 format. ECMA-376 doesn't strictly
	// mandate level order in document order, so we resolve level 0 by `w:ilvl`
	// rather than relying on `findChild` to return the first child as level 0.
	const targetFormat = kind === "bullet" ? "bullet" : "decimal";
	for (const child of root.findChildren("w:abstractNum")) {
		const lvl0 = findLevel(child, 0);
		const numFmt = lvl0?.findChild("w:numFmt");
		if (numFmt?.getAttribute("w:val") === targetFormat) {
			const id = child.getAttribute("w:abstractNumId");
			if (id) return Number(id);
		}
	}
	const newId = nextAbstractNumId(root);
	const def =
		kind === "bullet" ? bulletAbstractNum(newId) : orderedAbstractNum(newId);
	// Per OOXML, abstractNum elements come before num elements.
	const firstNumIdx = root.children.findIndex((child) => child.tag === "w:num");
	if (firstNumIdx === -1) root.children.push(def);
	else root.children.splice(firstNumIdx, 0, def);
	return newId;
}

function nextAbstractNumId(root: XmlNode): number {
	let max = -1;
	for (const child of root.findChildren("w:abstractNum")) {
		const id = child.getAttribute("w:abstractNumId");
		if (id) {
			const numeric = Number(id);
			if (Number.isFinite(numeric) && numeric > max) max = numeric;
		}
	}
	return max + 1;
}

function nextNumId(root: XmlNode): number {
	let max = 0;
	for (const child of root.findChildren("w:num")) {
		const id = child.getAttribute("w:numId");
		if (id) {
			const numeric = Number(id);
			if (Number.isFinite(numeric) && numeric > max) max = numeric;
		}
	}
	return max + 1;
}

/** Look up the `w:numFmt` for a specific (numId, level) pair by walking
 * `w:numbering` → `w:num` → `w:abstractNumId` → `w:abstractNum` → `w:lvl`.
 * Returns `undefined` if the doc has no numberingTree, the numId / level
 * isn't found, or the level has no numFmt. Falls back to level 0 if the
 * requested level isn't present (matches Word's render behavior of using
 * the lowest-defined level as the floor). */
export function getListFormat(
	view: DocView,
	numId: string,
	level: number,
): string | undefined {
	const tree = view.numberingTree;
	if (!tree) return undefined;
	const root = XmlNode.findRoot(tree, "w:numbering");
	if (!root) return undefined;
	const num = root
		.findChildren("w:num")
		.find((node) => node.getAttribute("w:numId") === numId);
	if (!num) return undefined;
	const abstractNumIdNode = num.findChild("w:abstractNumId");
	const abstractNumId = abstractNumIdNode?.getAttribute("w:val");
	if (!abstractNumId) return undefined;
	const abstractNum = root
		.findChildren("w:abstractNum")
		.find((node) => node.getAttribute("w:abstractNumId") === abstractNumId);
	if (!abstractNum) return undefined;
	const lvl = findLevel(abstractNum, level) ?? findLevel(abstractNum, 0);
	const numFmt = lvl?.findChild("w:numFmt");
	return numFmt?.getAttribute("w:val") ?? undefined;
}

function findLevel(abstractNum: XmlNode, ilvl: number): XmlNode | undefined {
	const target = String(ilvl);
	return abstractNum
		.findChildren("w:lvl")
		.find((lvl) => lvl.getAttribute("w:ilvl") === target);
}

/** Resolve the bullet glyph for a (numId, level) pair: `<w:lvl><w:lvlText w:val>`.
 *  Used by the GFM task-list reader to detect Word for Web's "Checklist"
 *  format (lvlText is the Wingdings ☐ glyph U+F0A8). Falls back to level 0 if
 *  the requested level is missing, mirroring `getListFormat`. */
export function getListBulletText(
	view: DocView,
	numId: string,
	level: number,
): string | undefined {
	const tree = view.numberingTree;
	if (!tree) return undefined;
	const root = XmlNode.findRoot(tree, "w:numbering");
	if (!root) return undefined;
	const num = root
		.findChildren("w:num")
		.find((node) => node.getAttribute("w:numId") === numId);
	if (!num) return undefined;
	const abstractNumId = num.findChild("w:abstractNumId")?.getAttribute("w:val");
	if (!abstractNumId) return undefined;
	const abstractNum = root
		.findChildren("w:abstractNum")
		.find((node) => node.getAttribute("w:abstractNumId") === abstractNumId);
	if (!abstractNum) return undefined;
	const lvl = findLevel(abstractNum, level) ?? findLevel(abstractNum, 0);
	return lvl?.findChild("w:lvlText")?.getAttribute("w:val") ?? undefined;
}

function bulletAbstractNum(abstractNumId: number): XmlNode {
	// Use Unicode bullet glyphs that render in any standard text font (Calibri
	// in our case via docDefaults). Earlier versions used Symbol/Wingdings as
	// the run font for the bullet glyph — those are symbol-substitution fonts,
	// so passing Unicode codepoints to them produces wrong/ugly glyphs (e.g.
	// "▪" through Wingdings renders as a malformed square).
	//
	// Levels 3-8 cycle the same three glyphs (•, ◦, ▪) — matches Word's
	// hybridMultilevel default. Agents emitting at level 3+ get a repeated
	// glyph rather than an undefined level.
	return (
		<w.abstractNum w-abstractNumId={String(abstractNumId)}>
			<w.multiLevelType w-val="hybridMultilevel" />
			{bulletLevel(0, "•")}
			{bulletLevel(1, "◦")}
			{bulletLevel(2, "▪")}
			{bulletLevel(3, "•")}
			{bulletLevel(4, "◦")}
			{bulletLevel(5, "▪")}
			{bulletLevel(6, "•")}
			{bulletLevel(7, "◦")}
			{bulletLevel(8, "▪")}
		</w.abstractNum>
	);
}

function bulletLevel(ilvl: number, bullet: string): XmlNode {
	return (
		<w.lvl w-ilvl={String(ilvl)}>
			<w.start w-val="1" />
			<w.numFmt w-val="bullet" />
			<w.lvlText w-val={bullet} />
			<w.lvlJc w-val="left" />
			<w.pPr>
				<w.ind w-left={String(levelIndent(ilvl))} w-hanging={HANGING} />
			</w.pPr>
		</w.lvl>
	);
}

function orderedAbstractNum(abstractNumId: number): XmlNode {
	return (
		<w.abstractNum w-abstractNumId={String(abstractNumId)}>
			<w.multiLevelType w-val="hybridMultilevel" />
			{orderedLevel(0, "%1.", "decimal")}
			{orderedLevel(1, "%2.", "lowerLetter")}
			{orderedLevel(2, "%3.", "lowerRoman")}
			{orderedLevel(3, "%4.", "decimal")}
			{orderedLevel(4, "%5.", "lowerLetter")}
			{orderedLevel(5, "%6.", "lowerRoman")}
			{orderedLevel(6, "%7.", "decimal")}
			{orderedLevel(7, "%8.", "lowerLetter")}
			{orderedLevel(8, "%9.", "lowerRoman")}
		</w.abstractNum>
	);
}

function orderedLevel(ilvl: number, text: string, fmt: string): XmlNode {
	return (
		<w.lvl w-ilvl={String(ilvl)}>
			<w.start w-val="1" />
			<w.numFmt w-val={fmt} />
			<w.lvlText w-val={text} />
			<w.lvlJc w-val="left" />
			<w.pPr>
				<w.ind w-left={String(levelIndent(ilvl))} w-hanging={HANGING} />
			</w.pPr>
		</w.lvl>
	);
}

/** Twips between paragraph-left and the first-line indent. Doubles as the
 * gap between bullet glyph and text — 240 twips ≈ 0.167" ≈ 16px at 96dpi,
 * which matches how most markdown readers (GitHub, Notion, Bear, etc.)
 * render lists. */
const HANGING = "240";

/** Twips of left indent per list level. Tighter than Word's
 * hybridMultilevel default (720 / 0.5" per level) — uses 300 / 0.208" per
 * level (≈ 20px at 96dpi), which feels about right next to typical markdown
 * rendering. Level 0's bullet sits a hair inside the margin (60 twips in)
 * with text at 0.208". */
function levelIndent(ilvl: number): number {
	return (ilvl + 1) * 300;
}
