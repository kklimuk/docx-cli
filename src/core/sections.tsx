import type { Document } from "./ast/document";
import type { SectionType } from "./ast/types";
import { w } from "./jsx";
import { XmlNode } from "./parser";
import type { TrackedMeta } from "./track-changes";

const EMU_PER_TWIP = 635;
/** US Letter (12240 twip) with 1″ (1440 twip) margins → 6.5″ content width. */
const DEFAULT_CONTENT_WIDTH_EMU = (12240 - 1440 - 1440) * EMU_PER_TWIP;

/** Page content width in EMU (page width − left/right margins), read from the
 *  document's trailing `<w:sectPr>` (`<w:pgSz>`/`<w:pgMar>`, in twips). Falls
 *  back to US-Letter-with-1″-margins when section properties are absent or
 *  malformed. Used to clamp inserted images so they don't spill into margins. */
export function getPageContentWidthEmu(document: Document): number {
	const root = XmlNode.findRoot(document.documentTree, "w:document");
	const body = root?.findChild("w:body");
	const sectPr = body?.children
		.filter((child) => child.tag === "w:sectPr")
		.pop();
	const width = Number(sectPr?.findChild("w:pgSz")?.getAttribute("w:w"));
	const pgMar = sectPr?.findChild("w:pgMar");
	const left = Number(pgMar?.getAttribute("w:left"));
	const right = Number(pgMar?.getAttribute("w:right"));
	if (
		!Number.isFinite(width) ||
		!Number.isFinite(left) ||
		!Number.isFinite(right)
	) {
		return DEFAULT_CONTENT_WIDTH_EMU;
	}
	const contentTwips = width - left - right;
	return contentTwips > 0
		? contentTwips * EMU_PER_TWIP
		: DEFAULT_CONTENT_WIDTH_EMU;
}

export type SectionProperties = {
	columns?: number;
	sectionType?: SectionType;
	pageWidth?: number;
	pageHeight?: number;
	pageOrientation?: "portrait" | "landscape";
	marginTop?: number;
	marginRight?: number;
	marginBottom?: number;
	marginLeft?: number;
};

/** Page-geometry intent for `applyPageGeometry` — the authorable subset of a
 *  `<w:sectPr>`. `pageSize` is PORTRAIT-normalized (`width <= height`); the final
 *  `<w:pgSz w:w/w:h>` is arranged by `orientation` (landscape swaps them). Each
 *  field is independent: pass only what changes. Margins are twips, signed (a
 *  negative margin pulls content into the header/footer area, which Word allows). */
export type PageGeometry = {
	pageSize?: { width: number; height: number };
	orientation?: "portrait" | "landscape";
	margins?: { top: number; right: number; bottom: number; left: number };
};

/** Extract columns / sectionType / page geometry from a list of sectPr children
 * (works for both the live sectPr's children and a sectPrChange snapshot's
 * inner sectPr children). Geometry (`<w:pgSz>`/`<w:pgMar>`) is read in twips. */
export function readSectionProperties(
	children: ReadonlyArray<XmlNode>,
): SectionProperties {
	const props: SectionProperties = {};
	for (const child of children) {
		if (child.tag === "w:cols") {
			const num = child.getAttribute("w:num");
			if (num) {
				const parsed = Number.parseInt(num, 10);
				if (Number.isFinite(parsed) && parsed > 0) props.columns = parsed;
			}
			continue;
		}
		if (child.tag === "w:type") {
			const value = child.getAttribute("w:val");
			if (value && isSectionType(value)) props.sectionType = value;
			continue;
		}
		if (child.tag === "w:pgSz") {
			const width = twipsAttr(child, "w:w");
			const height = twipsAttr(child, "w:h");
			if (width !== undefined) props.pageWidth = width;
			if (height !== undefined) props.pageHeight = height;
			const orient = child.getAttribute("w:orient");
			if (orient === "landscape" || orient === "portrait") {
				props.pageOrientation = orient;
			}
			continue;
		}
		if (child.tag === "w:pgMar") {
			// Margins are signed (a negative top is legal — content into the header
			// area). pgSz dimensions are positive.
			const top = twipsAttr(child, "w:top");
			const right = twipsAttr(child, "w:right");
			const bottom = twipsAttr(child, "w:bottom");
			const left = twipsAttr(child, "w:left");
			if (top !== undefined) props.marginTop = top;
			if (right !== undefined) props.marginRight = right;
			if (bottom !== undefined) props.marginBottom = bottom;
			if (left !== undefined) props.marginLeft = left;
		}
	}
	return props;
}

function twipsAttr(node: XmlNode, attr: string): number | undefined {
	const raw = node.getAttribute(attr);
	if (raw === undefined) return undefined;
	const parsed = Number.parseInt(raw, 10);
	return Number.isFinite(parsed) ? parsed : undefined;
}

const SECTION_TYPE_ORDER: ReadonlyArray<string> = [
	"continuous",
	"nextPage",
	"evenPage",
	"oddPage",
	"nextColumn",
];

export function isSectionType(value: string): value is SectionType {
	return SECTION_TYPE_ORDER.includes(value);
}

/** Emit a sentinel paragraph carrying an inline <w:sectPr> with the given
 * properties. Used by `docx sections`. The paragraph has no runs — its
 * sole purpose is to mark the end of a section. */
export function SentinelSectionParagraph({
	columns,
	sectionType,
}: {
	columns?: number;
	sectionType?: SectionType;
}): XmlNode {
	return (
		<w.p>
			<w.pPr>
				<w.sectPr>
					{sectionType && <w.type w-val={sectionType} />}
					{columns !== undefined && <w.cols w-num={String(columns)} />}
				</w.sectPr>
			</w.pPr>
		</w.p>
	);
}

/** Set or unset <w:cols w:num="N"/> inside an existing <w:sectPr>. Pass
 * `undefined` to leave existing columns untouched; pass `null` to clear. */
export function applyColumns(
	sectPr: XmlNode,
	columns: number | null | undefined,
): void {
	if (columns === undefined) return;
	const existing = sectPr.findChild("w:cols");
	if (columns === null) {
		if (existing) {
			const index = sectPr.children.indexOf(existing);
			if (index !== -1) sectPr.children.splice(index, 1);
		}
		return;
	}
	if (existing) {
		existing.setAttribute("w:num", String(columns));
		return;
	}
	insertSectPrChildInOrder(sectPr, <w.cols w-num={String(columns)} />);
}

/** Set or unset <w:type w:val="T"/> inside an existing <w:sectPr>. Pass
 * `undefined` to leave existing type untouched; pass `null` to clear. */
export function applySectionType(
	sectPr: XmlNode,
	sectionType: SectionType | null | undefined,
): void {
	if (sectionType === undefined) return;
	const existing = sectPr.findChild("w:type");
	if (sectionType === null) {
		if (existing) {
			const index = sectPr.children.indexOf(existing);
			if (index !== -1) sectPr.children.splice(index, 1);
		}
		return;
	}
	if (existing) {
		existing.setAttribute("w:val", sectionType);
		return;
	}
	insertSectPrChildInOrder(sectPr, <w.type w-val={sectionType} />);
}

/** Apply page geometry (size / orientation / margins) to an existing <w:sectPr>,
 * in place. Size + orientation are coupled (Word stores landscape as swapped
 * `<w:pgSz w:w/w:h>` PLUS `w:orient="landscape"`), so they're resolved together
 * from the current sectPr state + the requested deltas:
 *   - target size pair = requested `pageSize` (portrait-normalized) else the
 *     current pgSz (min/max) else US-Letter;
 *   - target orientation = requested `orientation` else the current one (the
 *     `w:orient` attr, or `w > h`) else portrait;
 *   - emit `w:w`/`w:h` as (short,long) for portrait or (long,short) for landscape,
 *     and set/clear `w:orient` accordingly (portrait is the default → attr dropped).
 * Margins are independent — each provided edge overwrites its `<w:pgMar>` attr.
 * Pass `undefined` fields to leave them untouched. CT_SectPr order is preserved
 * via `insertSectPrChildInOrder`. */
export function applyPageGeometry(
	sectPr: XmlNode,
	geometry: PageGeometry,
): void {
	if (geometry.pageSize || geometry.orientation) {
		const pgSz = findOrCreateSectPrChild(sectPr, "w:pgSz");
		const currentWidth = twipsAttr(pgSz, "w:w");
		const currentHeight = twipsAttr(pgSz, "w:h");
		const requested = geometry.pageSize;
		const short = requested
			? Math.min(requested.width, requested.height)
			: Math.min(currentWidth ?? 12240, currentHeight ?? 15840);
		const long = requested
			? Math.max(requested.width, requested.height)
			: Math.max(currentWidth ?? 12240, currentHeight ?? 15840);
		const currentOrient =
			pgSz.getAttribute("w:orient") ??
			((currentWidth ?? 0) > (currentHeight ?? 0) ? "landscape" : "portrait");
		const orientation = geometry.orientation ?? currentOrient;
		const landscape = orientation === "landscape";
		pgSz.setAttribute("w:w", String(landscape ? long : short));
		pgSz.setAttribute("w:h", String(landscape ? short : long));
		if (landscape) pgSz.setAttribute("w:orient", "landscape");
		else delete pgSz.attributes["w:orient"];
	}
	if (geometry.margins) {
		const pgMar = findOrCreateSectPrChild(sectPr, "w:pgMar");
		pgMar.setAttribute("w:top", String(geometry.margins.top));
		pgMar.setAttribute("w:right", String(geometry.margins.right));
		pgMar.setAttribute("w:bottom", String(geometry.margins.bottom));
		pgMar.setAttribute("w:left", String(geometry.margins.left));
	}
}

/** Copy a source sectPr's page geometry (`<w:pgSz>` size/orientation + `<w:pgMar>`
 *  margins) into `target`, but ONLY for the parts `target` doesn't already set.
 *  Page geometry is per-section in OOXML, so splitting a section (a column wrap
 *  minting fresh sentinel sectPrs) would otherwise drop the document's size/
 *  orientation/margins and silently revert those sections to the portrait-Letter
 *  default — the landscape-disappears-after-columns footgun. Clones the whole node
 *  (preserving attrs we don't model: header/footer/gutter). */
export function inheritPageGeometry(target: XmlNode, source: XmlNode): void {
	for (const tag of ["w:pgSz", "w:pgMar"] as const) {
		if (target.findChild(tag)) continue;
		const node = source.findChild(tag);
		if (node) insertSectPrChildInOrder(target, node.clone());
	}
}

/** Find a sectPr child by tag, or create + splice it at its CT_SectPr slot. */
function findOrCreateSectPrChild(sectPr: XmlNode, tag: string): XmlNode {
	const existing = sectPr.findChild(tag);
	if (existing) return existing;
	const created = new XmlNode(tag);
	insertSectPrChildInOrder(sectPr, created);
	return created;
}

// CT_SectPr child sequence (ECMA-376 §17.6.17), the subset we emit or step over.
// Word REJECTS an out-of-order sectPr ("unreadable content"); <w:sectPrChange>
// must be LAST (§17.6.18). Any code adding a child to an existing sectPr must
// splice via `insertSectPrChildInOrder`, never `push`.
const SECTPR_CHILD_ORDER = [
	"w:headerReference",
	"w:footerReference",
	"w:footnotePr",
	"w:endnotePr",
	"w:type",
	"w:pgSz",
	"w:pgMar",
	"w:paperSrc",
	"w:pgBorders",
	"w:lnNumType",
	"w:pgNumType",
	"w:cols",
	"w:formProt",
	"w:vAlign",
	"w:noEndnote",
	"w:titlePg",
	"w:textDirection",
	"w:bidi",
	"w:rtlGutter",
	"w:docGrid",
	"w:printerSettings",
	"w:sectPrChange",
] as const;

/** Rank a sectPr child by CT_SectPr position. Unknown tags rank just before
 *  `<w:sectPrChange>` so they still land ahead of the trailing change marker. */
function sectPrChildRank(tag: string): number {
	const index = SECTPR_CHILD_ORDER.indexOf(
		tag as (typeof SECTPR_CHILD_ORDER)[number],
	);
	if (index >= 0) return index;
	return SECTPR_CHILD_ORDER.indexOf("w:sectPrChange") - 0.5;
}

/** Splice `child` into `sectPr.children` at its canonical CT_SectPr position:
 *  before the first existing child that ranks after it (so a new `<w:pgSz>` lands
 *  after `<w:type>` but before `<w:cols>`, and everything before `<w:sectPrChange>`).
 *  Exported for the `Marginals` lens, which splices `<w:headerReference>` /
 *  `<w:footerReference>` / `<w:titlePg>` into existing sectPrs. */
export function insertSectPrChildInOrder(
	sectPr: XmlNode,
	child: XmlNode,
): void {
	const rank = sectPrChildRank(child.tag);
	const at = sectPr.children.findIndex(
		(existing) => sectPrChildRank(existing.tag) > rank,
	);
	if (at < 0) sectPr.children.push(child);
	else sectPr.children.splice(at, 0, child);
}

/** Delete an inline <w:sectPr> (lives inside a paragraph's <w:pPr>). The
 * owning paragraph stays — only the section break is removed. Returns true
 * if removed, false if the parent didn't contain the sectPr. */
export function removeInlineSectPr(
	sectPr: XmlNode,
	parentChildren: XmlNode[],
): boolean {
	const index = parentChildren.indexOf(sectPr);
	if (index === -1) return false;
	parentChildren.splice(index, 1);
	return true;
}

/** True if the given sectPr is the trailing one (a direct child of <w:body>),
 * which OOXML requires to exist. */
export function isTrailingSectPr(
	bodyChildren: XmlNode[],
	parentChildren: XmlNode[],
): boolean {
	return parentChildren === bodyChildren;
}

/** Snapshot the current section properties into a <w:sectPrChange> wrapper,
 * embed it in the sectPr, and return the change marker. The caller mutates
 * the live sectPr (via applyColumns / applySectionType) AFTER this call so
 * the snapshot captures the prior state.
 *
 * If a prior <w:sectPrChange> already exists it is replaced — we don't nest
 * multiple revisions; the most recent edit subsumes the prior. */
export function wrapSectPrChange(sectPr: XmlNode, meta: TrackedMeta): XmlNode {
	const existingChangeIndex = sectPr.children.findIndex(
		(child) => child.tag === "w:sectPrChange",
	);
	if (existingChangeIndex !== -1) {
		sectPr.children.splice(existingChangeIndex, 1);
	}

	const snapshotChildren: XmlNode[] = [];
	for (const child of sectPr.children) {
		if (child.tag === "w:sectPrChange") continue;
		snapshotChildren.push(child.clone());
	}

	const change = (
		<w.sectPrChange
			w-id={meta.revisionId}
			w-author={meta.author}
			w-date={meta.date}
		>
			<w.sectPr>{snapshotChildren}</w.sectPr>
		</w.sectPrChange>
	);
	sectPr.children.push(change);
	return change;
}
