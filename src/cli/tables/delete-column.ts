import { describeForms, TrackChanges } from "@core";
import { parseColumnAt } from "@core/locators";
import type { XmlNode } from "@core/parser";
import {
	buildGrid,
	cellAt,
	type Grid,
	markCellTracked,
	resolveTableNode,
} from "@core/table";
import {
	EXIT,
	fail,
	openOrFail,
	resolveTracked,
	respond,
	respondAck,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const AT_FORMS = describeForms(["tableColumn"], "                     ");

const HELP = `docx tables delete-column — delete a table column

Usage:
  docx tables delete-column FILE --at tN:cC [options]

Required:
  --at LOCATOR       Column to delete. Supports:
${AT_FORMS}
                     See \`docx info locators\`.

Optional:
  --author NAME      Author for tracked deletion (default: $DOCX_AUTHOR)
  -o, --output PATH  Write to PATH instead of overwriting FILE
  --dry-run          Print what would change; do not write
  -v, --verbose      Print the success ack JSON
  -h, --help         Show this help

When track-changes is on, each cell of the column is marked as a tracked
deletion (<w:tcPr><w:cellDel/>); the grid column is trimmed on accept.
Rejected if the column passes through a horizontal merge (unmerge first) or
is the table's only column (delete the table instead).

Output:
  Silent on success (exit 0). --verbose prints {ok:true, operation, path, table,
  column, tracked}. --dry-run prints the preview object (no ok field). Errors
  print {code, error, hint?} with a nonzero exit.

Examples:
  docx tables delete-column doc.docx --at t0:c1
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			author: { type: "string" },
			track: { type: "boolean" },
			...SAVE_FLAGS,
		},
		HELP,
	);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const at = parsed.values.at as string | undefined;
	if (!at) return fail("USAGE", "Missing --at tN:cC", HELP);
	const target = parseColumnAt(at);
	if (!target) {
		return fail(
			"INVALID_LOCATOR",
			`--at must be a column locator like t0:c2 (got ${at})`,
		);
	}

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const tableNode = resolveTableNode(document, target.tableId);
	if (!tableNode) {
		return fail("BLOCK_NOT_FOUND", `Table not found: ${target.tableId}`);
	}

	const grid = buildGrid(tableNode);
	if (target.col >= grid.colCount) {
		return fail(
			"BLOCK_NOT_FOUND",
			`Column ${target.col} not found in ${target.tableId} (has ${grid.colCount} columns)`,
		);
	}
	if (grid.colCount === 1) {
		return fail(
			"TABLE_STRUCTURE",
			"Cannot delete the only column",
			`Delete the whole table with \`docx delete ${target.tableId}\`.`,
		);
	}
	const spanningRow = findSpanningRow(grid, target.col);
	if (spanningRow !== null) {
		return fail(
			"TABLE_STRUCTURE",
			`Column ${target.col} passes through a horizontal merge in row ${spanningRow}`,
			"Unmerge the spanning cell first.",
		);
	}

	const tracking = resolveTracked(document, parsed.values.track);
	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			operation: "tables.delete-column",
			dryRun: true,
			path,
			table: target.tableId,
			column: target.col,
			tracked: tracking,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	const author = parsed.values.author as string | undefined;
	for (const row of grid.rows) {
		const cell = cellAt(row, target.col);
		if (!cell) continue;
		if (tracking) {
			markCellTracked(
				cell.node,
				"del",
				new TrackChanges(document).mintMeta(author),
			);
		} else {
			const index = row.node.children.indexOf(cell.node);
			if (index !== -1) row.node.children.splice(index, 1);
		}
	}
	if (!tracking && grid.tblGrid) {
		removeGridColumn(grid.tblGrid, target.col);
	}

	await document.save(outputPath);

	await respondAck({
		ok: true,
		operation: "tables.delete-column",
		path: outputPath ?? path,
		table: target.tableId,
		column: target.col,
		tracked: tracking,
	});
	return EXIT.OK;
}

/** The first row whose cell at logical `col` spans more than one column (so
 * deleting the column would split a horizontal merge), or null when safe. */
function findSpanningRow(grid: Grid, col: number): number | null {
	for (let index = 0; index < grid.rows.length; index++) {
		const row = grid.rows[index];
		if (!row) continue;
		const cell = cellAt(row, col);
		if (cell && cell.colSpan > 1) return index;
	}
	return null;
}

function removeGridColumn(tblGrid: XmlNode, col: number): void {
	const cols = tblGrid.findChildren("w:gridCol");
	const target = cols[col];
	if (!target) return;
	const index = tblGrid.children.indexOf(target);
	if (index !== -1) tblGrid.children.splice(index, 1);
}
