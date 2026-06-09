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
 * properties. Used by `insert --section`. The paragraph has no runs — its
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
	insertBeforeSectPrChange(sectPr, <w.cols w-num={String(columns)} />);
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
	const node = <w.type w-val={sectionType} />;
	const colsIndex = sectPr.children.findIndex(
		(child) => child.tag === "w:cols",
	);
	if (colsIndex !== -1) {
		sectPr.children.splice(colsIndex, 0, node);
		return;
	}
	insertBeforeSectPrChange(sectPr, node);
}

// Per ECMA-376 §17.6.18, <w:sectPrChange> must be the LAST child of <w:sectPr>.
// Append before any existing sectPrChange so the schema order survives a
// tracked edit that adds a new sectPr property.
function insertBeforeSectPrChange(sectPr: XmlNode, node: XmlNode): void {
	const changeIndex = sectPr.children.findIndex(
		(child) => child.tag === "w:sectPrChange",
	);
	if (changeIndex === -1) sectPr.children.push(node);
	else sectPr.children.splice(changeIndex, 0, node);
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
