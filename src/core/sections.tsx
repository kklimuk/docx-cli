import type { SectionType } from "./ast/types";
import { w } from "./jsx";
import type { XmlNode } from "./parser";
import type { TrackedMeta } from "./track-changes";

export type SectionProperties = {
	columns?: number;
	sectionType?: SectionType;
};

/** Extract the columns / sectionType pair from a list of sectPr children
 * (works for both the live sectPr's children and a sectPrChange snapshot's
 * inner sectPr children). */
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
		}
	}
	return props;
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
