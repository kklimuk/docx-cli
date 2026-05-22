import type { TableWidth } from "./ast/types";
import { Paragraph } from "./blocks";
import { type JsxChild, w } from "./jsx";
import type { NullableXmlNode, XmlNode } from "./parser";

/** Table emitters for `<w:tbl>` / `<w:tr>` / `<w:tc>`.
 *
 * `<BlankTable rows cols />` is the high-level entry point — it builds a
 * rows×cols grid of empty cells and is what `insert --table` (and the markdown
 * walker, for plain GFM tables) uses. For tables with cell content or merges,
 * compose the lower-level `<Table>` / `<TableRow>` / `<TableCell>` primitives
 * directly.
 *
 * Shared knobs (on `<BlankTable>` and `<Table>` alike):
 *   - `grid`: column widths in twips; length = logical column count.
 *   - `width`: total table width; defaults to 100% (5000 in OOXML pct units,
 *     which are fiftieths of a percent).
 *   - `borders`: "default" (single 0.5pt all-around + inside dividers) | "none"
 *     | a custom `{ style, sizeEighths, color }`.
 *   - `layout`: "autofit" (default — columns size to content, matches how GFM
 *     renders) | "fixed" (Word honors `grid` exactly; needed for custom widths
 *     or so merged cells align to the grid). */

/** A rows×cols table of empty cells with even-width columns. The high-level
 * entry point — used by `insert --table` and as a starting point for the
 * markdown walker's GFM tables. */
export function BlankTable({
	rows,
	cols,
	widths,
	width,
	borders,
	layout,
}: {
	rows: number;
	cols: number;
	/** Column widths in twips. Must have `cols` entries. If omitted, total
	 * table width is split evenly across columns. */
	widths?: number[];
	width?: TableWidth;
	borders?: TableBorders;
	layout?: TableLayout;
}): XmlNode {
	if (rows < 1) throw new Error(`Table must have at least 1 row (got ${rows})`);
	if (cols < 1) throw new Error(`Table must have at least 1 col (got ${cols})`);
	const grid = widths ?? evenGrid(cols);
	if (grid.length !== cols) {
		throw new Error(
			`Grid widths length (${grid.length}) doesn't match column count (${cols})`,
		);
	}
	return (
		<Table grid={grid} width={width} borders={borders} layout={layout}>
			{Array.from({ length: rows }, () => (
				<TableRow>
					{Array.from({ length: cols }, () => (
						<TableCell>
							<Paragraph text="" />
						</TableCell>
					))}
				</TableRow>
			))}
		</Table>
	);
}

/** Compose tables with arbitrary cell content and merges:
 *
 *   <Table grid={[2880, 2880]}>
 *     <TableRow>
 *       <TableCell><Paragraph text="A1" /></TableCell>
 *       <TableCell><Paragraph text="A2" /></TableCell>
 *     </TableRow>
 *     <TableRow>
 *       <TableCell gridSpan={2}><Paragraph text="spans both" /></TableCell>
 *     </TableRow>
 *   </Table>
 *
 * Cell merges: `gridSpan` spans N columns horizontally; `vMerge` ("restart" on
 * the top cell, "continue" below) merges down. */
export function Table({
	grid,
	width = { value: 5000, unit: "pct" },
	borders = "default",
	layout = "autofit",
	children,
}: {
	grid: number[];
	width?: TableWidth;
	borders?: TableBorders;
	layout?: TableLayout;
	children: JsxChild;
}): XmlNode {
	return (
		<w.tbl>
			<TableProperties width={width} borders={borders} layout={layout} />
			<TableGrid widths={grid} />
			{children}
		</w.tbl>
	);
}

export type TableLayout = "autofit" | "fixed";

export type TableBorderStyle = "single" | "double" | "none";

export type TableBorders =
	| "default"
	| "none"
	| {
			style: TableBorderStyle;
			/** Border thickness in eighths of a point (Word's `w:sz`). 4 = 0.5pt. */
			sizeEighths?: number;
			/** Hex color without `#` (e.g. "000000"). "auto" picks the theme color. */
			color?: string;
	  };

export function TableRow({ children }: { children: JsxChild }): XmlNode {
	return <w.tr>{children}</w.tr>;
}

export function TableCell({
	gridSpan,
	vMerge,
	width,
	children,
}: {
	gridSpan?: number;
	vMerge?: "restart" | "continue";
	width?: TableWidth;
	children: JsxChild;
}): XmlNode {
	return (
		<w.tc>
			<TableCellProperties gridSpan={gridSpan} vMerge={vMerge} width={width} />
			{children}
		</w.tc>
	);
}

function TableProperties({
	width,
	borders,
	layout,
}: {
	width: TableWidth;
	borders: TableBorders;
	layout: TableLayout;
}): XmlNode {
	// tblPr children follow CT_TblPr order (ECMA-376 §17.4.60):
	// tblW → tblBorders → tblLayout. Fixed layout makes Word honor `tblGrid`
	// exactly instead of recomputing column widths from cell content (the
	// autofit default), which otherwise flattens custom widths and misaligns
	// merged cells against the grid.
	return (
		<w.tblPr>
			<w.tblW w-w={String(width.value)} w-type={width.unit} />
			<TableBordersElement borders={borders} />
			<w.tblLayout w-type={layout} />
		</w.tblPr>
	);
}

function TableBordersElement({
	borders,
}: {
	borders: TableBorders;
}): NullableXmlNode {
	if (borders === "none") return null;
	const resolved =
		borders === "default"
			? { style: "single" as const, sizeEighths: 4, color: "auto" }
			: {
					style: borders.style,
					sizeEighths: borders.sizeEighths ?? 4,
					color: borders.color ?? "auto",
				};
	if (resolved.style === "none") return null;
	const attrs = {
		"w-val": resolved.style,
		"w-sz": String(resolved.sizeEighths),
		"w-space": "0",
		"w-color": resolved.color,
	};
	return (
		<w.tblBorders>
			<w.top {...attrs} />
			<w.left {...attrs} />
			<w.bottom {...attrs} />
			<w.right {...attrs} />
			<w.insideH {...attrs} />
			<w.insideV {...attrs} />
		</w.tblBorders>
	);
}

function TableGrid({ widths }: { widths: number[] }): XmlNode {
	return (
		<w.tblGrid>
			{widths.map((value) => (
				<w.gridCol w-w={String(value)} />
			))}
		</w.tblGrid>
	);
}

function TableCellProperties({
	gridSpan,
	vMerge,
	width,
}: {
	gridSpan: number | undefined;
	vMerge: "restart" | "continue" | undefined;
	width: TableWidth | undefined;
}): NullableXmlNode {
	if (!gridSpan && !vMerge && !width) return null;
	return (
		<w.tcPr>
			{width && <w.tcW w-w={String(width.value)} w-type={width.unit} />}
			{gridSpan !== undefined && gridSpan > 1 && (
				<w.gridSpan w-val={String(gridSpan)} />
			)}
			{vMerge && (
				<w.vMerge {...(vMerge === "restart" ? { "w-val": "restart" } : {})} />
			)}
		</w.tcPr>
	);
}

/** Default total table width: 9360 twips (6.5") — letter-paper page width
 * minus the default 1" margins. Splits evenly. */
function evenGrid(cols: number): number[] {
	const total = 9360;
	const base = Math.floor(total / cols);
	const remainder = total - base * cols;
	return Array.from(
		{ length: cols },
		(_, i) => base + (i === 0 ? remainder : 0),
	);
}
