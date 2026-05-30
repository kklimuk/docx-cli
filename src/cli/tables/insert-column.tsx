import { TrackChanges } from "@core";
import { parseTableAt } from "@core/locators";
import type { XmlNode } from "@core/parser";
import {
	appendTblGridChange,
	buildGrid,
	emptyCell,
	type Grid,
	type GridRow,
	gridColElement,
	markCellTracked,
	resolveTableNode,
} from "@core/table";
import { parseArgs } from "util";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	setVerboseAck,
	writeStdout,
} from "../respond";

const HELP = `docx tables insert-column — insert a table column

Usage:
  docx tables insert-column FILE --at tN [options]

Required:
  --at tN            Target table (e.g. t0)

Optional:
  --position INDEX   0-based column index to insert at (default: append at end)
  --width TWIPS      Width of the new column in twips (default: average of the
                     existing columns)
  --author NAME      Author for tracked insertion (default: $DOCX_AUTHOR)
  -o, --output PATH  Write to PATH instead of overwriting FILE
  --dry-run          Print what would change; do not write
  -v, --verbose      Print the success ack JSON
  -h, --help         Show this help

When track-changes is on, each new cell is wrapped as a tracked insertion
(<w:tcPr><w:cellIns/>). Rejected if the insertion point bisects a horizontal
merge — unmerge first.

Examples:
  docx tables insert-column doc.docx --at t0
  docx tables insert-column doc.docx --at t0 --position 1 --width 1440
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				at: { type: "string" },
				position: { type: "string" },
				width: { type: "string" },
				author: { type: "string" },
				output: { type: "string", short: "o" },
				"dry-run": { type: "boolean" },
				verbose: { type: "boolean", short: "v" },
				help: { type: "boolean", short: "h" },
			},
		});
	} catch (parseError) {
		const message =
			parseError instanceof Error ? parseError.message : String(parseError);
		return fail("USAGE", message, HELP);
	}

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const at = parsed.values.at as string | undefined;
	if (!at) return fail("USAGE", "Missing --at tN", HELP);
	const tableId = parseTableAt(at);
	if (!tableId) {
		return fail(
			"INVALID_LOCATOR",
			`--at must be a table id like t0 (got ${at})`,
		);
	}

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const tableNode = resolveTableNode(document, tableId);
	if (!tableNode) return fail("BLOCK_NOT_FOUND", `Table not found: ${tableId}`);

	const grid = buildGrid(tableNode);
	if (!grid.tblGrid) {
		return fail("TABLE_STRUCTURE", `Table ${tableId} has no <w:tblGrid>`);
	}

	const position = resolveColumnPosition(parsed.values.position, grid.colCount);
	if (position === null) {
		return fail(
			"USAGE",
			`--position must be an integer in 0..${grid.colCount}`,
		);
	}

	const bisected = findBisectedRow(grid, position);
	if (bisected !== null) {
		return fail(
			"TABLE_STRUCTURE",
			`Inserting a column at ${position} would bisect a horizontal merge in row ${bisected}`,
			"Unmerge the spanning cell first.",
		);
	}

	const existingWidths = grid.tblGrid
		.findChildren("w:gridCol")
		.map((col) => Number(col.getAttribute("w:w") ?? "0"));
	const width = resolveWidth(parsed.values.width, existingWidths);
	if (width === null)
		return fail("USAGE", "--width must be a positive integer");

	const tracking = document.isTrackChangesEnabled();
	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "tables.insert-column",
			dryRun: true,
			path,
			table: tableId,
			position,
			width,
			tracked: tracking,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	const author = parsed.values.author as string | undefined;
	const priorCols = grid.tblGrid.findChildren("w:gridCol");
	insertGridColumn(grid.tblGrid, position, width);
	if (tracking) {
		// Pair the per-cell cellIns marks with a grid-revision snapshot so the
		// width change is reversible (Word keeps the grown grid; reject restores
		// the prior one).
		appendTblGridChange(
			grid.tblGrid,
			priorCols,
			new TrackChanges(document).mintMeta(author),
		);
	}
	for (const row of grid.rows) {
		const cell = emptyCell();
		if (tracking) {
			markCellTracked(cell, "ins", new TrackChanges(document).mintMeta(author));
		}
		insertCellAtColumn(row, position, cell);
	}

	await document.save(outputPath);

	await respondAck({
		ok: true,
		operation: "tables.insert-column",
		path: outputPath ?? path,
		table: tableId,
		position,
		width,
		tracked: tracking,
	});
	return EXIT.OK;
}

function resolveColumnPosition(raw: unknown, colCount: number): number | null {
	if (raw === undefined) return colCount;
	const value = Number(raw);
	if (!Number.isInteger(value) || value < 0 || value > colCount) return null;
	return value;
}

function resolveWidth(raw: unknown, existing: number[]): number | null {
	if (raw === undefined) {
		const positive = existing.filter((value) => value > 0);
		if (positive.length === 0) return 1440;
		return Math.round(
			positive.reduce((sum, v) => sum + v, 0) / positive.length,
		);
	}
	const value = Number(raw);
	if (!Number.isInteger(value) || value <= 0) return null;
	return value;
}

/** The first row whose cell strictly spans across the boundary at `position`
 * (so the inserted column would split a merged cell), or null when safe. */
function findBisectedRow(grid: Grid, position: number): number | null {
	for (let index = 0; index < grid.rows.length; index++) {
		const row = grid.rows[index];
		if (!row) continue;
		const bisects = row.cells.some(
			(cell) =>
				cell.colStart < position && cell.colStart + cell.colSpan > position,
		);
		if (bisects) return index;
	}
	return null;
}

function insertGridColumn(
	tblGrid: XmlNode,
	position: number,
	width: number,
): void {
	const cols = tblGrid.findChildren("w:gridCol");
	const anchor = cols[position];
	const at = anchor
		? tblGrid.children.indexOf(anchor)
		: tblGrid.children.length;
	tblGrid.children.splice(at, 0, gridColElement(width));
}

function insertCellAtColumn(
	row: GridRow,
	position: number,
	cell: XmlNode,
): void {
	const anchorCell = row.cells.find((entry) => entry.colStart === position);
	if (anchorCell) {
		const at = row.node.children.indexOf(anchorCell.node);
		row.node.children.splice(at, 0, cell);
		return;
	}
	// Appending past the last cell: place after the final <w:tc>.
	const lastCell = row.cells[row.cells.length - 1];
	const at = lastCell
		? row.node.children.indexOf(lastCell.node) + 1
		: row.node.children.length;
	row.node.children.splice(at, 0, cell);
}
