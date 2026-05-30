import { w } from "../../jsx";
import { XmlNode } from "../../parser";
import type { ContentTypesView } from "./content-types";
import type { Pkg } from "./package";
import type { RelationshipsView } from "./relationships";

export type AbstractNumKind = "bullet" | "ordered";

const NUMBERING_PART_NAME = "word/numbering.xml";
const NUMBERING_RELATIONSHIP_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
const NUMBERING_CONTENT_TYPE =
	"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";

const EMPTY_NUMBERING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;

export class NumberingView {
	tree: XmlNode[];

	constructor(tree: XmlNode[] = XmlNode.parse(EMPTY_NUMBERING_XML)) {
		this.tree = tree;
	}

	/** Load this view from a package; returns undefined if the part is absent. */
	static async fromPackage(pkg: Pkg): Promise<NumberingView | undefined> {
		const tree = await pkg.readPart(NUMBERING_PART_NAME);
		return tree ? new NumberingView(tree) : undefined;
	}

	/** Parse a view from raw XML; returns undefined if the input is absent. */
	static fromXml(xml: string | undefined): NumberingView | undefined {
		return xml ? new NumberingView(XmlNode.parse(xml)) : undefined;
	}

	/** Serialize this view's tree into the package's `word/numbering.xml`. */
	writeTo(pkg: Pkg): void {
		pkg.writeText(NUMBERING_PART_NAME, XmlNode.serialize(this.tree));
	}

	/** Mint the numbering relationship + content-type override on the
	 * containing package and return a fresh empty view. Idempotent on the
	 * relationship target. Called by `Document.ensureNumbering()`. */
	static register(deps: {
		relationships: RelationshipsView;
		contentTypes: ContentTypesView;
	}): NumberingView {
		if (!deps.relationships.hasTarget("numbering.xml")) {
			deps.relationships.add(NUMBERING_RELATIONSHIP_TYPE, "numbering.xml");
		}
		deps.contentTypes.registerPart(NUMBERING_PART_NAME, NUMBERING_CONTENT_TYPE);
		return new NumberingView();
	}

	listNumIds(): string[] {
		const root = XmlNode.findRoot(this.tree, "w:numbering");
		if (!root) return [];
		const out: string[] = [];
		for (const child of root.findChildren("w:num")) {
			const id = child.getAttribute("w:numId");
			if (id) out.push(id);
		}
		return out;
	}

	listAbstractNumIds(): string[] {
		const root = XmlNode.findRoot(this.tree, "w:numbering");
		if (!root) return [];
		const out: string[] = [];
		for (const child of root.findChildren("w:abstractNum")) {
			const id = child.getAttribute("w:abstractNumId");
			if (id) out.push(id);
		}
		return out;
	}

	/** Allocate a fresh `numId` for a new list of the given kind, pointing at
	 * the package's shared bullet or ordered `abstractNum`. Each call returns a
	 * distinct `numId` so multiple lists in the same document restart their
	 * numbering independently (per OOXML convention: abstractNum is the style;
	 * num is the per-list instance). Seeds the bullet/ordered `abstractNum` on
	 * first use. Style provisioning for the `ListParagraph` pStyle is the
	 * caller's responsibility — typically via
	 * `view.ensureStyles().ensureStyle("ListParagraph")`. */
	allocate(kind: AbstractNumKind): number {
		const root = this.ensureNumberingRoot();
		const abstractNumId = this.ensureAbstractNum(root, kind);
		const numId = nextNumId(root);
		root.children.push(
			<NumElement numId={numId} abstractNumId={abstractNumId} />,
		);
		return numId;
	}

	/** Look up the `w:numFmt` for a specific (numId, level) pair by walking
	 * `w:numbering` → `w:num` → `w:abstractNumId` → `w:abstractNum` → `w:lvl`.
	 * Returns `undefined` if the numId / level isn't found, or the level has no
	 * numFmt. Falls back to level 0 if the requested level isn't present
	 * (matches Word's render behavior of using the lowest-defined level as the
	 * floor). */
	getFormat(numId: string, level: number): string | undefined {
		const abstractNum = this.resolveAbstractNum(numId);
		if (!abstractNum) return undefined;
		const lvl = findLevel(abstractNum, level) ?? findLevel(abstractNum, 0);
		return lvl?.findChild("w:numFmt")?.getAttribute("w:val") ?? undefined;
	}

	/** Resolve the bullet glyph for a (numId, level) pair:
	 * `<w:lvl><w:lvlText w:val>`. Used by the GFM task-list reader to detect
	 * Word for Web's "Checklist" format (lvlText is the Wingdings ☐ glyph
	 * U+F0A8). Falls back to level 0 if the requested level is missing,
	 * mirroring `getFormat`. */
	getBulletText(numId: string, level: number): string | undefined {
		const abstractNum = this.resolveAbstractNum(numId);
		if (!abstractNum) return undefined;
		const lvl = findLevel(abstractNum, level) ?? findLevel(abstractNum, 0);
		return lvl?.findChild("w:lvlText")?.getAttribute("w:val") ?? undefined;
	}

	private resolveAbstractNum(numId: string): XmlNode | undefined {
		const root = XmlNode.findRoot(this.tree, "w:numbering");
		if (!root) return undefined;
		const num = root
			.findChildren("w:num")
			.find((node) => node.getAttribute("w:numId") === numId);
		if (!num) return undefined;
		const abstractNumId = num
			.findChild("w:abstractNumId")
			?.getAttribute("w:val");
		if (!abstractNumId) return undefined;
		return root
			.findChildren("w:abstractNum")
			.find((node) => node.getAttribute("w:abstractNumId") === abstractNumId);
	}

	private ensureNumberingRoot(): XmlNode {
		const root = XmlNode.findRoot(this.tree, "w:numbering");
		if (!root) {
			throw new Error("expected <w:numbering> root in numbering tree");
		}
		return root;
	}

	/** Reuse an existing abstractNum of the matching kind, or build + insert a
	 * fresh one. Reuse keys on level-0 numFmt alone — sufficient for our own
	 * abstractNums, but won't deduplicate identically-shaped abstractNums
	 * seeded by other tools that happen to share a level-0 format. */
	private ensureAbstractNum(root: XmlNode, kind: AbstractNumKind): number {
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
		const def = <AbstractNum kind={kind} id={newId} />;
		// Per OOXML, abstractNum elements come before num elements.
		const firstNumIdx = root.children.findIndex(
			(child) => child.tag === "w:num",
		);
		if (firstNumIdx === -1) root.children.push(def);
		else root.children.splice(firstNumIdx, 0, def);
		return newId;
	}
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

function findLevel(abstractNum: XmlNode, ilvl: number): XmlNode | undefined {
	const target = String(ilvl);
	return abstractNum
		.findChildren("w:lvl")
		.find((lvl) => lvl.getAttribute("w:ilvl") === target);
}

/** `<w:num w:numId="N"><w:abstractNumId w:val="K"/></w:num>` — the per-list
 * instance pointer at a shared abstractNum. */
function NumElement({
	numId,
	abstractNumId,
}: {
	numId: number;
	abstractNumId: number;
}): XmlNode {
	return (
		<w.num w-numId={String(numId)}>
			<w.abstractNumId w-val={String(abstractNumId)} />
		</w.num>
	);
}

/** Build a fresh `<w:abstractNum>` of the given kind at the given id. Bullet
 * uses Unicode `•/◦/▪` glyphs that render in any standard text font (avoids
 * the symbol-substitution issue with Wingdings); ordered uses the
 * `decimal/lowerLetter/lowerRoman` hybrid Word emits by default. */
function AbstractNum({
	kind,
	id,
}: {
	kind: AbstractNumKind;
	id: number;
}): XmlNode {
	return kind === "bullet" ? (
		<BulletAbstractNum id={id} />
	) : (
		<OrderedAbstractNum id={id} />
	);
}

/** Bullet `<w:abstractNum>` with all nine levels populated. Levels 3-8 cycle
 * the same three glyphs (•, ◦, ▪) — matches Word's hybridMultilevel default
 * so agents emitting at level 3+ get a repeated glyph rather than an undefined
 * level. */
function BulletAbstractNum({ id }: { id: number }): XmlNode {
	return (
		<w.abstractNum w-abstractNumId={String(id)}>
			<w.multiLevelType w-val="hybridMultilevel" />
			<BulletLevel ilvl={0} glyph="•" />
			<BulletLevel ilvl={1} glyph="◦" />
			<BulletLevel ilvl={2} glyph="▪" />
			<BulletLevel ilvl={3} glyph="•" />
			<BulletLevel ilvl={4} glyph="◦" />
			<BulletLevel ilvl={5} glyph="▪" />
			<BulletLevel ilvl={6} glyph="•" />
			<BulletLevel ilvl={7} glyph="◦" />
			<BulletLevel ilvl={8} glyph="▪" />
		</w.abstractNum>
	);
}

function BulletLevel({
	ilvl,
	glyph,
}: {
	ilvl: number;
	glyph: string;
}): XmlNode {
	return (
		<w.lvl w-ilvl={String(ilvl)}>
			<w.start w-val="1" />
			<w.numFmt w-val="bullet" />
			<w.lvlText w-val={glyph} />
			<w.lvlJc w-val="left" />
			<w.pPr>
				<w.ind w-left={String(levelIndent(ilvl))} w-hanging={HANGING} />
			</w.pPr>
		</w.lvl>
	);
}

function OrderedAbstractNum({ id }: { id: number }): XmlNode {
	return (
		<w.abstractNum w-abstractNumId={String(id)}>
			<w.multiLevelType w-val="hybridMultilevel" />
			<OrderedLevel ilvl={0} text="%1." fmt="decimal" />
			<OrderedLevel ilvl={1} text="%2." fmt="lowerLetter" />
			<OrderedLevel ilvl={2} text="%3." fmt="lowerRoman" />
			<OrderedLevel ilvl={3} text="%4." fmt="decimal" />
			<OrderedLevel ilvl={4} text="%5." fmt="lowerLetter" />
			<OrderedLevel ilvl={5} text="%6." fmt="lowerRoman" />
			<OrderedLevel ilvl={6} text="%7." fmt="decimal" />
			<OrderedLevel ilvl={7} text="%8." fmt="lowerLetter" />
			<OrderedLevel ilvl={8} text="%9." fmt="lowerRoman" />
		</w.abstractNum>
	);
}

function OrderedLevel({
	ilvl,
	text,
	fmt,
}: {
	ilvl: number;
	text: string;
	fmt: string;
}): XmlNode {
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
