import { w } from "../../jsx";
import { XmlNode } from "../../parser";
import type { ContentTypesView } from "./content-types";
import type { Pkg } from "./package";
import type { RelationshipsView } from "./relationships";

export type AbstractNumKind = "bullet" | "ordered";

/** The `docx lists set --format` vocabulary for ordered lists. Friendly names
 * map to OOXML `<w:numFmt>` values; the AST/markdown hint carries the friendly
 * form so an agent can read a value and feed it straight back to `--format`. */
export type ListFormat =
	| "decimal"
	| "lower-alpha"
	| "upper-alpha"
	| "lower-roman"
	| "upper-roman";

export const FORMAT_TO_NUMFMT: Record<ListFormat, string> = {
	decimal: "decimal",
	"lower-alpha": "lowerLetter",
	"upper-alpha": "upperLetter",
	"lower-roman": "lowerRoman",
	"upper-roman": "upperRoman",
};

const NUMFMT_TO_FORMAT: Record<string, ListFormat> = {
	decimal: "decimal",
	lowerLetter: "lower-alpha",
	upperLetter: "upper-alpha",
	lowerRoman: "lower-roman",
	upperRoman: "upper-roman",
};

/** Map an OOXML `<w:numFmt>` value to the friendly `--format` vocabulary when
 * it's one we model; pass exotic formats through untouched. */
export function numFmtToFormat(numFmt: string): string {
	return NUMFMT_TO_FORMAT[numFmt] ?? numFmt;
}

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
	 * `view.ensureStyles().ensureStyle("ListParagraph")`.
	 *
	 * Pass `start` to set the level-0 start value via `<w:lvlOverride>` —
	 * needed both to honor `1. / 10. ordered` lists from markdown and to
	 * stop Word from auto-continuing adjacent lists into one. Defaults to
	 * 1, which is the safe restart-from-the-top behavior. */
	allocate(kind: AbstractNumKind, start = 1): number {
		const root = this.ensureNumberingRoot();
		const abstractNumId = this.ensureAbstractNum(root, kind);
		const numId = nextNumId(root);
		root.children.push(
			<NumElement numId={numId} abstractNumId={abstractNumId} start={start} />,
		);
		return numId;
	}

	/** One-pass resolution of the (numId, level) numbering facts `read` needs for
	 * every list paragraph: the effective numFmt (`format` — override or
	 * abstractNum), the AUTHORED override numFmt alone (`override`), and the
	 * effective `start`. Resolving the num + abstractNum once here keeps the read
	 * path from re-walking the numbering tree per getter for each list item. Falls
	 * back to level 0 when the requested level is absent (Word's lowest-defined
	 * floor). `getFormat`/`getStart` are thin views over it. */
	getLevelInfo(
		numId: string,
		level: number,
	): { format?: string; override?: string; start: number } {
		const num = this.findNum(numId);
		const lvlOverride = num ? this.findLvlOverride(num, level) : undefined;
		const override = lvlOverride
			?.findChild("w:lvl")
			?.findChild("w:numFmt")
			?.getAttribute("w:val");
		const abstractNum = this.resolveAbstractNum(numId);
		const abstractLvl = abstractNum
			? (findLevel(abstractNum, level) ?? findLevel(abstractNum, 0))
			: undefined;
		const format =
			override ?? abstractLvl?.findChild("w:numFmt")?.getAttribute("w:val");
		const startRaw =
			lvlOverride?.findChild("w:startOverride")?.getAttribute("w:val") ??
			abstractLvl?.findChild("w:start")?.getAttribute("w:val");
		// A malformed (non-numeric) start in a foreign doc must not poison the
		// markdown ordinal with NaN — fall back to 1. A legitimate 0 (a 0-based
		// list) is finite and kept.
		const startNum = Number(startRaw);
		const start =
			startRaw !== undefined && Number.isFinite(startNum) ? startNum : 1;
		return { format, override, start };
	}

	/** The effective `w:numFmt` for a (numId, level): per-list override wins, then
	 * the abstractNum (level-0 floor). `undefined` if neither is found. (The
	 * AUTHORED override alone is `getLevelInfo(...).override`, which `read` uses to
	 * populate `Paragraph.list.format`.) */
	getFormat(numId: string, level: number): string | undefined {
		return this.getLevelInfo(numId, level).format;
	}

	/** Resolve the bullet glyph for a (numId, level) pair:
	 * `<w:lvl><w:lvlText w:val>`. Used by the GFM task-list reader to detect
	 * Word for Web's "Checklist" format (lvlText is the Wingdings ☐ glyph
	 * U+F0A8). A per-list `<w:lvlOverride>` wins, then the abstractNum; falls
	 * back to level 0 if the requested level is missing, mirroring `getFormat`. */
	getBulletText(numId: string, level: number): string | undefined {
		const fromOverride = this.overrideLevel(numId, level)
			?.findChild("w:lvlText")
			?.getAttribute("w:val");
		if (fromOverride !== undefined) return fromOverride;
		const abstractNum = this.resolveAbstractNum(numId);
		if (!abstractNum) return undefined;
		const lvl = findLevel(abstractNum, level) ?? findLevel(abstractNum, 0);
		return lvl?.findChild("w:lvlText")?.getAttribute("w:val") ?? undefined;
	}

	/** The effective level-`level` start value for list `numId`: a per-list
	 * `<w:lvlOverride><w:startOverride>` wins, then the abstractNum lvl's
	 * `<w:start>`, else 1. `read` surfaces this as `Paragraph.list.start` so a
	 * `--start N` list round-trips through the markdown ordinal. */
	getStart(numId: string, level: number): number {
		return this.getLevelInfo(numId, level).start;
	}

	/** Set (or update) the level-`level` start value for list `numId` via its
	 * `<w:lvlOverride><w:startOverride>`. Returns false if the numId is absent.
	 * Affects every paragraph carrying `numId` — the list as a whole. */
	setStart(numId: string, level: number, start: number): boolean {
		const num = this.findNum(numId);
		if (!num) return false;
		const lvlOverride = this.ensureLvlOverride(num, level);
		const existing = lvlOverride.findChild("w:startOverride");
		if (existing) existing.setAttribute("w:val", String(start));
		else
			insertLvlOverrideChildInOrder(
				lvlOverride,
				<w.startOverride w-val={String(start)} />,
			);
		return true;
	}

	/** Override the numbering format for level `level` of list `numId`, per-list,
	 * via `<w:lvlOverride><w:lvl>`. Word rejects a partial `<w:lvl>`, so clone the
	 * backing abstractNum's level (a well-formed CT_Lvl) and swap only its
	 * `<w:numFmt>` — the `%N.` lvlText already renders the formatted counter for
	 * any numFmt. Returns false if the numId is absent. */
	setFormat(numId: string, level: number, numFmt: string): boolean {
		const num = this.findNum(numId);
		if (!num) return false;
		const lvl = this.buildOverrideLevel(numId, level, numFmt);
		const lvlOverride = this.ensureLvlOverride(num, level);
		const existing = lvlOverride.findChild("w:lvl");
		if (existing) {
			lvlOverride.children.splice(
				lvlOverride.children.indexOf(existing),
				1,
				lvl,
			);
		} else {
			insertLvlOverrideChildInOrder(lvlOverride, lvl);
		}
		return true;
	}

	/** Mint a fresh numId reproducing `srcNumId`'s abstractNum and any per-list
	 * format override, but (re)starting level 0 at `start`. Used by `--restart`
	 * to split one list into an independently-numbered one. */
	cloneListDefinition(srcNumId: string, start: number): number {
		const root = this.ensureNumberingRoot();
		const src = this.findNum(srcNumId);
		const abstractNumId =
			src?.findChild("w:abstractNumId")?.getAttribute("w:val") ?? "0";
		const numId = nextNumId(root);
		const num = <w.num w-numId={String(numId)} />;
		num.children.push(<w.abstractNumId w-val={abstractNumId} />);
		// Carry forward any format overrides so a restarted list keeps its glyphs.
		if (src) {
			for (const lvlOverride of src.findChildren("w:lvlOverride")) {
				const lvl = lvlOverride.findChild("w:lvl");
				if (!lvl) continue;
				const carried = (
					<w.lvlOverride w-ilvl={lvlOverride.getAttribute("w:ilvl") ?? "0"} />
				);
				carried.children.push(lvl.clone());
				num.children.push(carried);
			}
		}
		// Push before `setStart`, which resolves the num by id from the tree.
		root.children.push(num);
		this.setStart(String(numId), 0, start);
		return numId;
	}

	private findNum(numId: string): XmlNode | undefined {
		const root = XmlNode.findRoot(this.tree, "w:numbering");
		return root
			?.findChildren("w:num")
			.find((node) => node.getAttribute("w:numId") === numId);
	}

	private findLvlOverride(num: XmlNode, level: number): XmlNode | undefined {
		const target = String(level);
		return num
			.findChildren("w:lvlOverride")
			.find((node) => node.getAttribute("w:ilvl") === target);
	}

	/** The `<w:num>`'s level-`level` `<w:lvlOverride>`, creating it (in CT_Num
	 * order — after `<w:abstractNumId>`) if absent. */
	private ensureLvlOverride(num: XmlNode, level: number): XmlNode {
		const existing = this.findLvlOverride(num, level);
		if (existing) return existing;
		const created = <w.lvlOverride w-ilvl={String(level)} />;
		const abstractIdx = num.children.findIndex(
			(child) => child.tag === "w:abstractNumId",
		);
		num.children.splice(abstractIdx + 1, 0, created);
		return created;
	}

	/** The per-list `<w:lvlOverride><w:lvl>` override for (numId, level), if any —
	 * the source of a format/glyph override `getFormat`/`getBulletText` consult. */
	private overrideLevel(numId: string, level: number): XmlNode | undefined {
		const num = this.findNum(numId);
		const lvlOverride = num ? this.findLvlOverride(num, level) : undefined;
		return lvlOverride?.findChild("w:lvl");
	}

	/** A well-formed CT_Lvl for an override: clone the abstractNum's level (so we
	 * keep its indent pPr / lvlText) and swap the numFmt. Falls back to a fresh
	 * ordered level if the abstractNum can't be resolved. Drops the cloned
	 * `<w:start>` so the sibling `<w:startOverride>` is the SOLE source of the
	 * level's start — keeping both (with possibly different values) is internally
	 * contradictory and risks a reader/renderer picking the stale inner start. */
	private buildOverrideLevel(
		numId: string,
		level: number,
		numFmt: string,
	): XmlNode {
		const abstractNum = this.resolveAbstractNum(numId);
		const base =
			abstractNum &&
			(findLevel(abstractNum, level) ?? findLevel(abstractNum, 0));
		const lvl = base ? (
			base.clone()
		) : (
			<OrderedLevel ilvl={level} text={`%${level + 1}.`} fmt={numFmt} />
		);
		lvl.setAttribute("w:ilvl", String(level));
		lvl.findChild("w:numFmt")?.setAttribute("w:val", numFmt);
		lvl.children = lvl.children.filter((child) => child.tag !== "w:start");
		return lvl;
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

/** CT_LvlOverride is an ordered sequence (ECMA-376 §17.9.8): `<w:startOverride>`
 * then `<w:lvl>`. Splice a child into that order — so a `setStart` after a
 * `setFormat` lands its `<w:startOverride>` BEFORE the existing `<w:lvl>`. The
 * numbering analog of `insertPprChildInOrder`. */
function insertLvlOverrideChildInOrder(
	lvlOverride: XmlNode,
	child: XmlNode,
): void {
	const order = ["w:startOverride", "w:lvl"];
	const rank = order.indexOf(child.tag);
	const at = lvlOverride.children.findIndex(
		(existing) => !existing.isText && order.indexOf(existing.tag) > rank,
	);
	if (at < 0) lvlOverride.children.push(child);
	else lvlOverride.children.splice(at, 0, child);
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
	start,
}: {
	numId: number;
	abstractNumId: number;
	start: number;
}): XmlNode {
	// `<w:lvlOverride>` + `<w:startOverride>` forces this num instance to
	// (re)start at the given value. Always emitting it (even for start=1)
	// stops Word's "continue numbering" autoformat from merging adjacent
	// lists that happen to share the abstractNum, and lets `start=10` from
	// `10. item` survive into Word's rendered output.
	return (
		<w.num w-numId={String(numId)}>
			<w.abstractNumId w-val={String(abstractNumId)} />
			<w.lvlOverride w-ilvl="0">
				<w.startOverride w-val={String(start)} />
			</w.lvlOverride>
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
