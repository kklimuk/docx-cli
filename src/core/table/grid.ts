import type { DocView } from "../ast/doc-view";
import type { XmlNode } from "../parser";

/** A logical view of a `<w:tbl>` that resolves `gridSpan` / `vMerge` into a
 * coordinate system. Each physical `<w:tc>` occupies `colSpan` consecutive
 * logical columns starting at `colStart`; a `vMerge="continue"` cell is still
 * a real `<w:tc>` in its row. The `tables` CLI verbs query the model to map
 * logical (row, col) coordinates onto the physical cells they mutate. */

/** Resolve a `tN` (or chained `tN:rRcC:tK…`) locator to its `<w:tbl>` element,
 * or null when the id is absent from `view.blockReferences` or points at a
 * non-table node. */
export function resolveTableNode(
	view: DocView,
	tableId: string,
): XmlNode | null {
	const reference = view.blockReferences.get(tableId);
	if (!reference || reference.node.tag !== "w:tbl") return null;
	return reference.node;
}

export type GridCell = {
	/** The physical `<w:tc>` element. */
	node: XmlNode;
	/** Logical column this cell begins at (0-based). */
	colStart: number;
	/** Number of logical columns this cell spans (`w:gridSpan`, ≥ 1). */
	colSpan: number;
	vMerge?: "restart" | "continue";
};

export type GridRow = {
	/** The physical `<w:tr>` element. */
	node: XmlNode;
	cells: GridCell[];
};

export type Grid = {
	table: XmlNode;
	tblGrid: XmlNode | undefined;
	rows: GridRow[];
	/** Logical column count — `<w:tblGrid>` length, falling back to the widest
	 * row when the grid is missing or out of sync. */
	colCount: number;
};

export function buildGrid(table: XmlNode): Grid {
	const tblGrid = table.findChild("w:tblGrid");
	const gridColCount = tblGrid ? tblGrid.findChildren("w:gridCol").length : 0;
	const rows: GridRow[] = [];
	let widestRow = 0;
	for (const rowNode of table.findChildren("w:tr")) {
		const cells: GridCell[] = [];
		let col = 0;
		for (const cellNode of rowNode.findChildren("w:tc")) {
			const colSpan = readGridSpan(cellNode);
			const vMerge = readVMerge(cellNode);
			const cell: GridCell = { node: cellNode, colStart: col, colSpan };
			if (vMerge) cell.vMerge = vMerge;
			cells.push(cell);
			col += colSpan;
		}
		widestRow = Math.max(widestRow, col);
		rows.push({ node: rowNode, cells });
	}
	return { table, tblGrid, rows, colCount: gridColCount || widestRow };
}

/** The physical cell whose logical span covers `col` in this row, or undefined
 * when the row is absent or `col` is past the row's last cell. */
export function cellAt(
	row: GridRow | undefined,
	col: number,
): GridCell | undefined {
	return row?.cells.find(
		(cell) => cell.colStart <= col && col < cell.colStart + cell.colSpan,
	);
}

export function readGridSpan(cell: XmlNode): number {
	const raw = cell
		.findChild("w:tcPr")
		?.findChild("w:gridSpan")
		?.getAttribute("w:val");
	const value = raw ? Number(raw) : Number.NaN;
	return Number.isFinite(value) && value > 1 ? value : 1;
}

export function readVMerge(cell: XmlNode): "restart" | "continue" | undefined {
	const marker = cell.findChild("w:tcPr")?.findChild("w:vMerge");
	if (!marker) return undefined;
	// ECMA-376 §17.4.84: w:val="restart" begins a merge; a bare element (or
	// "continue") continues it from the cell above.
	return marker.getAttribute("w:val") === "restart" ? "restart" : "continue";
}
