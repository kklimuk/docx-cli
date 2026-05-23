import { saveDocView } from "@core";
import { parseCellRangeAt } from "@core/locators";
import {
	buildGrid,
	cellAt,
	clearCellContent,
	type Grid,
	type GridRow,
	resolveTableNode,
	setGridSpan,
	setVMerge,
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
import { noteStructuralChange } from "./common";

const HELP = `docx tables merge — merge a rectangular cell region

Usage:
  docx tables merge FILE --at tN:rR1cC1-rR2cC2 [options]

Required:
  --at tN:rR1cC1-rR2cC2   Top-left to bottom-right cell region (e.g. t0:r0c0-r1c1)

Optional:
  --author NAME      Author for the audit comment when track-changes is on
  -o, --output PATH  Write to PATH instead of overwriting FILE
  --dry-run          Print what would change; do not write
  -v, --verbose      Print the success ack JSON
  -h, --help         Show this help

Horizontal extent collapses into the leftmost cell via <w:gridSpan>; vertical
extent uses <w:vMerge> ("restart" on the top row, "continue" below). The
top-left cell keeps its content; the rest are emptied. OOXML has no
tracked-change construct for merges, so under track-changes the merge applies
immediately with a [docx-cli] audit comment.

Rejected if the region edges bisect an existing merge — unmerge first.

Examples:
  docx tables merge doc.docx --at t0:r0c0-r0c2     # merge 3 cells across
  docx tables merge doc.docx --at t0:r0c0-r2c0     # merge 3 cells down
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				at: { type: "string" },
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
	if (!at) return fail("USAGE", "Missing --at tN:rR1cC1-rR2cC2", HELP);
	const region = parseCellRangeAt(at);
	if (!region) {
		return fail(
			"INVALID_LOCATOR",
			`--at must be a cell-range locator like t0:r0c0-r1c1 (got ${at})`,
		);
	}

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const tableNode = resolveTableNode(view, region.tableId);
	if (!tableNode) {
		return fail("BLOCK_NOT_FOUND", `Table not found: ${region.tableId}`);
	}

	const grid = buildGrid(tableNode);
	const r1 = Math.min(region.start.row, region.end.row);
	const r2 = Math.max(region.start.row, region.end.row);
	const c1 = Math.min(region.start.col, region.end.col);
	const c2 = Math.max(region.start.col, region.end.col);

	if (r1 === r2 && c1 === c2) {
		return fail("USAGE", "Merge region must cover more than one cell");
	}
	if (r2 >= grid.rows.length || c2 >= grid.colCount) {
		return fail(
			"BLOCK_NOT_FOUND",
			`Region ${at} is out of bounds (table is ${grid.rows.length}×${grid.colCount})`,
		);
	}
	const misaligned = findMisalignedRow(grid, r1, r2, c1, c2);
	if (misaligned !== null) {
		return fail(
			"TABLE_STRUCTURE",
			`Region edges cross an existing merge in row ${misaligned}`,
			"Unmerge the spanning cell first.",
		);
	}

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "tables.merge",
			dryRun: true,
			path,
			table: region.tableId,
			region: { r1, c1, r2, c2 },
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	const width = c2 - c1 + 1;
	for (let rowIndex = r1; rowIndex <= r2; rowIndex++) {
		const row = grid.rows[rowIndex];
		if (!row) continue;
		mergeRow(row, c1, c2, width, rowIndex === r1, r1 !== r2);
	}

	// Word applies cell merges immediately even under tracking (no revision
	// marker), so we do the same and note it with an audit comment.
	noteStructuralChange(
		view,
		cellAt(grid.rows[r1] as GridRow, c1)?.node,
		`cells merged (r${r1}c${c1}-r${r2}c${c2})`,
		parsed.values.author as string | undefined,
	);

	await saveDocView(view, outputPath);

	await respondAck({
		ok: true,
		operation: "tables.merge",
		path: outputPath ?? path,
		table: region.tableId,
		region: { r1, c1, r2, c2 },
	});
	return EXIT.OK;
}

/** The first row in [r1..r2] whose cells don't start exactly at c1 and end
 * exactly at c2 (so the region would split a merge), or null when aligned. */
function findMisalignedRow(
	grid: Grid,
	r1: number,
	r2: number,
	c1: number,
	c2: number,
): number | null {
	for (let rowIndex = r1; rowIndex <= r2; rowIndex++) {
		const row = grid.rows[rowIndex];
		if (!row) return rowIndex;
		const left = cellAt(row, c1);
		const right = cellAt(row, c2);
		if (!left || !right) return rowIndex;
		if (left.colStart !== c1) return rowIndex;
		if (right.colStart + right.colSpan - 1 !== c2) return rowIndex;
	}
	return null;
}

function mergeRow(
	row: GridRow,
	c1: number,
	c2: number,
	width: number,
	isTop: boolean,
	vertical: boolean,
): void {
	const anchor = cellAt(row, c1);
	if (!anchor) return;
	// Remove the physical cells to the right of the anchor within the region.
	for (const cell of row.cells) {
		if (cell === anchor) continue;
		if (cell.colStart < c1 || cell.colStart > c2) continue;
		const index = row.node.children.indexOf(cell.node);
		if (index !== -1) row.node.children.splice(index, 1);
	}
	if (width > 1) setGridSpan(anchor.node, width);
	if (vertical) setVMerge(anchor.node, isTop ? "restart" : "continue");
	if (vertical && !isTop) clearCellContent(anchor.node);
}
