import { CELL_MARGIN_SIDES, CELL_MARGIN_TAGS } from "../ast/cell-margins";
import type { CellMargins, TableWidth } from "../ast/types";
import { isInheritableRunProperty, Paragraph } from "../blocks";
import { w } from "../jsx";
import { XmlNode } from "../parser";
import type { TrackedMeta } from "../track-changes";
import { Del, Ins } from "../track-changes/emit";
import { TableCell } from ".";

/** Fresh-XML constructors and `<w:tcPr>` surgery shared by the table verbs.
 * The setters mutate an existing cell in place — they add/remove the
 * `gridSpan` / `vMerge` / `tcW` child without disturbing other `<w:tcPr>`
 * children (borders, shading, vAlign…) so unmodeled cell properties survive,
 * per the in-place-mutation invariant. */

/** An empty cell: `<w:tc>` wrapping a single empty paragraph (a `<w:tc>` must
 * contain at least one block-level child). */
export function emptyCell(): XmlNode {
	return (
		<TableCell>
			<Paragraph text="" />
		</TableCell>
	);
}

export function gridColElement(width: number): XmlNode {
	return <w.gridCol w-w={String(width)} />;
}

/** Replace a cell's block content with a single empty paragraph — used when a
 * cell becomes a `vMerge` continuation (Word renders the merge anchor's content
 * and ignores the continued cells'). */
export function clearCellContent(cell: XmlNode): void {
	const tcPr = cell.findChild("w:tcPr");
	cell.children = tcPr ? [tcPr] : [];
	cell.children.push(<Paragraph text="" />);
}

export function setGridSpan(cell: XmlNode, span: number): void {
	const tcPr = ensureTcPr(cell);
	setTcPrChild(
		tcPr,
		"w:gridSpan",
		span > 1 ? <w.gridSpan w-val={String(span)} /> : null,
	);
	pruneEmptyTcPr(cell, tcPr);
}

export function setVMerge(
	cell: XmlNode,
	value: "restart" | "continue" | null,
): void {
	const tcPr = ensureTcPr(cell);
	const node =
		value === "restart" ? (
			<w.vMerge w-val="restart" />
		) : value === "continue" ? (
			<w.vMerge />
		) : null;
	setTcPrChild(tcPr, "w:vMerge", node);
	pruneEmptyTcPr(cell, tcPr);
}

export function setCellWidth(cell: XmlNode, width: TableWidth | null): void {
	const tcPr = ensureTcPr(cell);
	setTcPrChild(
		tcPr,
		"w:tcW",
		width ? <w.tcW w-w={String(width.value)} w-type={width.unit} /> : null,
	);
	pruneEmptyTcPr(cell, tcPr);
}

/** Set or clear a cell's background fill via `<w:tcPr><w:shd w:fill="…"/>`.
 * `hex` is a 6-digit color without `#`; `null` removes the shading. */
export function setCellShading(cell: XmlNode, hex: string | null): void {
	const tcPr = ensureTcPr(cell);
	setTcPrChild(
		tcPr,
		"w:shd",
		hex ? <w.shd w-val="clear" w-color="auto" w-fill={hex} /> : null,
	);
	pruneEmptyTcPr(cell, tcPr);
}

/** Set or clear a cell's vertical alignment via `<w:tcPr><w:vAlign w:val="…"/>`.
 * `null` (or `"top"`, the default) removes the element. */
export function setCellVAlign(
	cell: XmlNode,
	value: "top" | "center" | "bottom" | null,
): void {
	const tcPr = ensureTcPr(cell);
	setTcPrChild(
		tcPr,
		"w:vAlign",
		value && value !== "top" ? <w.vAlign w-val={value} /> : null,
	);
	pruneEmptyTcPr(cell, tcPr);
}

/** Clone the reference (sibling-row) cell's `<w:tcBorders>` and `<w:tcMar>`
 * onto a freshly built cell, so an inserted row keeps the table's rule lines and
 * cell padding — without the borders the new row renders with no separators
 * while every sibling row has them (the invoice "Calibration kit" defect,
 * flagged by two independent judge runs); without the padding it renders at the
 * table default while its siblings carry the template's own (mnda's 7.2pt rows).
 * Shading and vAlign are deliberately NOT inherited: banded/striped rows make
 * those unsafe to assume, while borders and padding are uniform across data rows
 * in practice. Each part is a no-op when the reference lacks it or the cell
 * already has its own. */
export function inheritCellBordersAndMargins(
	cell: XmlNode,
	referenceCell: XmlNode,
): void {
	const referenceTcPr = referenceCell.findChild("w:tcPr");
	for (const tag of ["w:tcBorders", "w:tcMar"]) {
		const reference = referenceTcPr?.findChild(tag);
		if (!reference) continue;
		const tcPr = ensureTcPr(cell);
		if (tcPr.findChild(tag)) continue;
		setTcPrChild(tcPr, tag, reference.clone());
	}
}

/** Build a cell paragraph that inherits the reference cell's look — the
 * paragraph properties (alignment, spacing, indent) and the first run's
 * `<w:rPr>` — so a minted cell lines up with its column siblings. Without the
 * pPr an inserted numeric column falls back to the left-aligned default and
 * sits visibly out of column with the right-aligned values above it (the
 * invoice "Calibration kit" row); without the rPr its text falls back to
 * docDefaults and renders in a visibly different font/size. `<w:pPr>` carries
 * only properties (no content), so cloning it wholesale is safe; `<Paragraph>`
 * with no options emits no `<w:pPr>`, so the clone becomes the paragraph's
 * sole (and correctly first) `<w:pPr>`. The borders half of the same policy is
 * `inheritCellBordersAndMargins` above — a caller minting a full cell applies both. */
export function cellParagraphLike(
	referenceCell: XmlNode | undefined,
	text: string,
): XmlNode {
	const paragraph = <Paragraph text={text} />;
	const referenceParagraph = referenceCell?.findChild("w:p");
	const pPr = referenceParagraph?.findChild("w:pPr");
	if (pPr) paragraph.children.unshift(pPr.clone());
	const referenceRpr = referenceParagraph?.findChild("w:r")?.findChild("w:rPr");
	if (referenceRpr) {
		const inherited = referenceRpr.clone();
		// A fresh cell inherits the reference's FONT, not its placeholder-fill or
		// tracked-revision markers — the shared `isInheritableRunProperty` policy,
		// the same exclusion the body-edit `inheritCommonRunFormatting` applies.
		inherited.children = inherited.children.filter(isInheritableRunProperty);
		if (inherited.children.length > 0) {
			for (const run of paragraph.findChildren("w:r")) {
				run.children.unshift(inherited.clone());
			}
		}
	}
	return paragraph;
}

/** A single border edge's properties, shared by table and cell border builders. */
export type BorderEdge = { style: string; sizeEighths: number; color: string };

/** Merge per-side border edges into a cell's `<w:tcBorders>`. Each entry sets
 * (`edge`) or removes (`null`) one side, preserving sides not named — so
 * `--cell-borders bottom` adds a bottom rule without disturbing existing edges.
 * A `clearAll` request drops the whole `<w:tcBorders>`. Side children are kept in
 * CT_TcBorders order (ECMA-376 §17.4.66). */
export function setCellBorders(
	cell: XmlNode,
	updates: { side: CellBorderSide; edge: BorderEdge | null }[],
	clearAll = false,
): void {
	const tcPr = ensureTcPr(cell);
	if (clearAll) {
		setTcPrChild(tcPr, "w:tcBorders", null);
		pruneEmptyTcPr(cell, tcPr);
		return;
	}
	let tcBorders = tcPr.findChild("w:tcBorders");
	if (!tcBorders) {
		if (updates.every((update) => update.edge === null)) {
			pruneEmptyTcPr(cell, tcPr);
			return;
		}
		tcBorders = <w.tcBorders />;
		setTcPrChild(tcPr, "w:tcBorders", tcBorders);
	}
	for (const { side, edge } of updates) {
		const tag = `w:${side}`;
		tcBorders.children = tcBorders.children.filter(
			(child) => child.tag !== tag,
		);
		if (edge) insertEdgeInOrder(tcBorders, tag, borderEdge(tag, edge));
	}
	if (tcBorders.children.length === 0) setTcPrChild(tcPr, "w:tcBorders", null);
	pruneEmptyTcPr(cell, tcPr);
}

/** Build one `<w:top/>`/`<w:bottom/>`/… border edge element (dynamic tag, so
 * built via `XmlNode.element` rather than the static-tag JSX namespace). A `none`
 * style emits just `w:val="none"` (an explicit no-border that overrides
 * inheritance); any other style carries size/space/color. */
function borderEdge(tag: string, edge: BorderEdge): XmlNode {
	const attributes: Record<string, string> =
		edge.style === "none"
			? { "w:val": "none" }
			: {
					"w:val": edge.style,
					"w:sz": String(edge.sizeEighths),
					"w:space": "0",
					"w:color": edge.color,
				};
	return XmlNode.element(tag, attributes);
}

const CELL_BORDER_ORDER = [
	"w:top",
	"w:left",
	"w:bottom",
	"w:right",
	"w:insideH",
	"w:insideV",
];

function insertEdgeInOrder(parent: XmlNode, tag: string, node: XmlNode): void {
	const target = orderIndex(CELL_BORDER_ORDER, tag);
	let insertAt = parent.children.length;
	for (let index = 0; index < parent.children.length; index++) {
		const child = parent.children[index];
		if (child && orderIndex(CELL_BORDER_ORDER, child.tag) > target) {
			insertAt = index;
			break;
		}
	}
	parent.children.splice(insertAt, 0, node);
}

export type CellBorderSide =
	| "top"
	| "left"
	| "bottom"
	| "right"
	| "insideH"
	| "insideV";

/** Set or clear a cell's own padding via `<w:tcPr><w:tcMar>` (overrides the
 * table default). Only the sides `margins` names change — an absent side keeps
 * whatever the cell already had (so `-,4pt,-,4pt` touches left/right only).
 * `null` removes the element. */
export function setCellMargins(
	cell: XmlNode,
	margins: CellMargins | null,
): void {
	const tcPr = ensureTcPr(cell);
	if (!margins) setTcPrChild(tcPr, "w:tcMar", null);
	else {
		const existing = tcPr.findChild("w:tcMar");
		if (existing) mergeMarginSides(existing, margins);
		else setTcPrChild(tcPr, "w:tcMar", marginsElement("w:tcMar", margins));
	}
	pruneEmptyTcPr(cell, tcPr);
}

/** Set or clear the table's default cell padding via `<w:tblPr><w:tblCellMar>`
 * (Word's Table Properties → Options → Default cell margins). Like
 * `setCellMargins`, only the named sides change; `null` removes it. */
export function setTableCellMargins(
	table: XmlNode,
	margins: CellMargins | null,
): void {
	const existing = table.findChild("w:tblPr")?.findChild("w:tblCellMar");
	if (margins && existing) {
		mergeMarginSides(existing, margins);
		return;
	}
	setTablePropertiesChild(
		table,
		"w:tblCellMar",
		margins ? marginsElement("w:tblCellMar", margins) : null,
	);
}

function marginsElement(tag: string, margins: CellMargins): XmlNode {
	const node = XmlNode.element(tag, {});
	mergeMarginSides(node, margins);
	return node;
}

/** Write each side `margins` names onto a `<w:tblCellMar>`/`<w:tcMar>` in place:
 * an existing child for that side (`left` OR the strict `start`, …) is updated
 * under its own tag, a missing one is inserted in schema order as the
 * transitional `left`/`right` Word itself writes. Sides `margins` omits are
 * left untouched. */
function mergeMarginSides(node: XmlNode, margins: CellMargins): void {
	for (const side of CELL_MARGIN_SIDES) {
		const twips = margins[side];
		if (twips === undefined) continue;
		const attributes = { "w:w": String(twips), "w:type": "dxa" };
		const existing = CELL_MARGIN_TAGS.filter(([padded]) => padded === side)
			.map(([, tag]) => node.findChild(tag))
			.find((child) => child !== undefined);
		if (existing) {
			for (const [name, value] of Object.entries(attributes))
				existing.setAttribute(name, value);
			continue;
		}
		const tag = `w:${side}`;
		const rank = schemaRank(tag);
		const before = node.children.findIndex(
			(child) => schemaRank(child.tag) > rank,
		);
		const element = XmlNode.element(tag, attributes);
		if (before === -1) node.children.push(element);
		else node.children.splice(before, 0, element);
	}
}

const schemaRank = (tag: string): number =>
	CELL_MARGIN_TAGS.findIndex(([, candidate]) => candidate === tag);

/** Mark a whole row as a tracked insertion/deletion via `<w:trPr><w:ins/>` or
 * `<w:del/>` (ECMA-376 §17.13.5.18 / §17.13.5.13). */
export function markRowTracked(
	row: XmlNode,
	kind: "ins" | "del",
	meta: TrackedMeta,
): void {
	const trPr = ensureTrPr(row);
	trPr.children = trPr.children.filter(
		(child) => child.tag !== "w:ins" && child.tag !== "w:del",
	);
	// CT_TrPr places ins/del near the end (only trPrChange follows, which we
	// never emit), so appending is schema-correct.
	trPr.children.push(
		kind === "ins" ? <Ins meta={meta} /> : <Del meta={meta} />,
	);
}

/** Mark a single cell as a tracked insertion/deletion via `<w:tcPr><w:cellIns/>`
 * or `<w:cellDel/>` (ECMA-376 §17.4.16 / §17.4.10) — the per-cell record of a
 * tracked column insert/delete. */
export function markCellTracked(
	cell: XmlNode,
	kind: "ins" | "del",
	meta: TrackedMeta,
): void {
	const tcPr = ensureTcPr(cell);
	const attrs = {
		"w-id": String(meta.revisionId),
		"w-author": meta.author,
		"w-date": meta.date,
	};
	setTcPrChild(
		tcPr,
		kind === "ins" ? "w:cellIns" : "w:cellDel",
		kind === "ins" ? <w.cellIns {...attrs} /> : <w.cellDel {...attrs} />,
	);
}

function ensureTrPr(row: XmlNode): XmlNode {
	const existing = row.findChild("w:trPr");
	if (existing) return existing;
	const trPr: XmlNode = <w.trPr />;
	row.children.unshift(trPr);
	return trPr;
}

/** Set or clear a row's height via `<w:trPr><w:trHeight w:val w:hRule/>`.
 * `value` is in twips; `null` removes the element (height reverts to auto). */
export function setRowHeight(
	row: XmlNode,
	height: { value: number; rule: "atLeast" | "exact" | "auto" } | null,
): void {
	const trPr = ensureTrPr(row);
	setTrPrChild(
		trPr,
		"w:trHeight",
		height ? (
			<w.trHeight w-val={String(height.value)} w-hRule={height.rule} />
		) : null,
	);
	pruneEmptyTrPr(row, trPr);
}

/** Mark or unmark a row as a repeating header via `<w:trPr><w:tblHeader/>`. */
export function setRepeatHeader(row: XmlNode, on: boolean): void {
	const trPr = ensureTrPr(row);
	setTrPrChild(trPr, "w:tblHeader", on ? <w.tblHeader /> : null);
	pruneEmptyTrPr(row, trPr);
}

function pruneEmptyTrPr(row: XmlNode, trPr: XmlNode): void {
	if (trPr.children.length === 0) {
		row.children = row.children.filter((child) => child !== trPr);
	}
}

/** Append a `<w:tblGridChange>` snapshot of the prior grid columns to a
 * `<w:tblGrid>` whose `<w:gridCol>` set has just been mutated — the tracked
 * record of a column-width/count revision (ECMA-376 §17.4.49). `priorCols`
 * are clones of the gridCol elements as they were *before* the mutation.
 * Mirrors `wrapSectPrChange`. CT_TblGridChange extends CT_Markup, which
 * carries ONLY `w:id` — no author/date (the per-cell markers carry those). */
export function appendTblGridChange(
	tblGrid: XmlNode,
	priorCols: XmlNode[],
	meta: TrackedMeta,
): void {
	tblGrid.children = tblGrid.children.filter(
		(child) => child.tag !== "w:tblGridChange",
	);
	tblGrid.children.push(
		<w.tblGridChange w-id={String(meta.revisionId)}>
			<w.tblGrid>{priorCols.map((col) => col.clone())}</w.tblGrid>
		</w.tblGridChange>,
	);
}

/** Append a `<w:tcPrChange>` snapshot of the prior `<w:tcPr>` children to a
 * cell's `<w:tcPr>` (ECMA-376 §17.13.5.36) — the tracked record of a
 * cell-property revision (e.g. a `<w:tcW>` width change). `priorChildren` are
 * clones taken *before* the mutation. Mirrors `wrapSectPrChange`. */
export function appendTcPrChange(
	cell: XmlNode,
	priorChildren: XmlNode[],
	meta: TrackedMeta,
): void {
	const tcPr = ensureTcPr(cell);
	setTcPrChild(
		tcPr,
		"w:tcPrChange",
		<w.tcPrChange
			w-id={String(meta.revisionId)}
			w-author={meta.author}
			w-date={meta.date}
		>
			<w.tcPr>{priorChildren.map((child) => child.clone())}</w.tcPr>
		</w.tcPrChange>,
	);
}

function ensureTcPr(cell: XmlNode): XmlNode {
	const existing = cell.findChild("w:tcPr");
	if (existing) return existing;
	const tcPr: XmlNode = <w.tcPr />;
	cell.children.unshift(tcPr);
	return tcPr;
}

function pruneEmptyTcPr(cell: XmlNode, tcPr: XmlNode): void {
	if (tcPr.children.length === 0) {
		cell.children = cell.children.filter((child) => child !== tcPr);
	}
}

/** Find or create the `<w:tblPr>` (first child of `<w:tbl>`). */
export function ensureTblPr(table: XmlNode): XmlNode {
	const existing = table.findChild("w:tblPr");
	if (existing) return existing;
	const tblPr: XmlNode = <w.tblPr />;
	table.children.unshift(tblPr);
	return tblPr;
}

/** Set or replace `<w:tblLayout>` in a table's `<w:tblPr>`. */
export function setTableLayout(
	table: XmlNode,
	layout: "fixed" | "autofit",
): void {
	setTblPrChild(
		ensureTblPr(table),
		"w:tblLayout",
		<w.tblLayout w-type={layout} />,
	);
}

/** Set or remove an arbitrary `<w:tblPr>` child in schema order — exposed so
 * the borders verb can splice `<w:tblBorders>` in correctly. */
export function setTablePropertiesChild(
	table: XmlNode,
	tag: string,
	node: XmlNode | null,
): void {
	setTblPrChild(ensureTblPr(table), tag, node);
}

/** Set or clear the table's on-page justification via `<w:tblPr><w:jc w:val/>`.
 * `null` (or `"left"`, the default) removes the element. */
export function setTableJustification(
	table: XmlNode,
	value: "left" | "center" | "right" | null,
): void {
	setTablePropertiesChild(
		table,
		"w:jc",
		value && value !== "left" ? <w.jc w-val={value} /> : null,
	);
}

/** Set or clear the table's style reference via `<w:tblPr><w:tblStyle w:val/>`. */
export function setTableStyle(table: XmlNode, styleId: string | null): void {
	setTablePropertiesChild(
		table,
		"w:tblStyle",
		styleId ? <w.tblStyle w-val={styleId} /> : null,
	);
}

const TBL_PR_ORDER = [
	"w:tblStyle",
	"w:tblpPr",
	"w:tblOverlap",
	"w:bidiVisual",
	"w:tblStyleRowBandSize",
	"w:tblStyleColBandSize",
	"w:tblW",
	"w:jc",
	"w:tblCellSpacing",
	"w:tblInd",
	"w:tblBorders",
	"w:shd",
	"w:tblLayout",
	"w:tblCellMar",
	"w:tblLook",
	"w:tblPrChange",
];

function setTblPrChild(
	tblPr: XmlNode,
	tag: string,
	node: XmlNode | null,
): void {
	tblPr.children = tblPr.children.filter((child) => child.tag !== tag);
	if (!node) return;
	const target = orderIndex(TBL_PR_ORDER, tag);
	let insertAt = tblPr.children.length;
	for (let index = 0; index < tblPr.children.length; index++) {
		const child = tblPr.children[index];
		if (child && orderIndex(TBL_PR_ORDER, child.tag) > target) {
			insertAt = index;
			break;
		}
	}
	tblPr.children.splice(insertAt, 0, node);
}

function orderIndex(order: string[], tag: string): number {
	const index = order.indexOf(tag);
	return index === -1 ? order.length : index;
}

// CT_TcPr child order (ECMA-376 §17.4.42) — the subset we touch plus the
// neighbors we must order against.
const TC_PR_ORDER = [
	"w:cnfStyle",
	"w:tcW",
	"w:gridSpan",
	"w:hMerge",
	"w:vMerge",
	"w:tcBorders",
	"w:shd",
	"w:noWrap",
	"w:tcMar",
	"w:textDirection",
	"w:tcFitText",
	"w:vAlign",
	"w:hideMark",
	"w:cellIns",
	"w:cellDel",
	"w:cellMerge",
	"w:tcPrChange",
];

/** Remove any existing `<tag>` from `tcPr`, then (if `node` is given) splice it
 * back in at the schema-mandated position. */
function setTcPrChild(tcPr: XmlNode, tag: string, node: XmlNode | null): void {
	tcPr.children = tcPr.children.filter((child) => child.tag !== tag);
	if (!node) return;
	const target = orderIndex(TC_PR_ORDER, tag);
	let insertAt = tcPr.children.length;
	for (let index = 0; index < tcPr.children.length; index++) {
		const child = tcPr.children[index];
		if (child && orderIndex(TC_PR_ORDER, child.tag) > target) {
			insertAt = index;
			break;
		}
	}
	tcPr.children.splice(insertAt, 0, node);
}

// CT_TrPr child order (ECMA-376 §17.4.81) — the subset we touch plus the
// neighbors (ins/del/trPrChange come last) we must order against.
const TR_PR_ORDER = [
	"w:cnfStyle",
	"w:divId",
	"w:gridBefore",
	"w:gridAfter",
	"w:wBefore",
	"w:wAfter",
	"w:cantSplit",
	"w:trHeight",
	"w:tblHeader",
	"w:tblCellSpacing",
	"w:jc",
	"w:hidden",
	"w:ins",
	"w:del",
	"w:trPrChange",
];

/** Remove any existing `<tag>` from `trPr`, then (if `node` is given) splice it
 * back in at the schema-mandated position. */
function setTrPrChild(trPr: XmlNode, tag: string, node: XmlNode | null): void {
	trPr.children = trPr.children.filter((child) => child.tag !== tag);
	if (!node) return;
	const target = orderIndex(TR_PR_ORDER, tag);
	let insertAt = trPr.children.length;
	for (let index = 0; index < trPr.children.length; index++) {
		const child = trPr.children[index];
		if (child && orderIndex(TR_PR_ORDER, child.tag) > target) {
			insertAt = index;
			break;
		}
	}
	trPr.children.splice(insertAt, 0, node);
}
