import { describeForms, TrackChanges } from "@core";
import { Paragraph } from "@core/blocks";
import { w } from "@core/jsx";
import { parseTableAt } from "@core/locators";
import type { XmlNode } from "@core/parser";
import {
	buildGrid,
	cellAt,
	type Grid,
	type GridRow,
	markRowTracked,
	resolveTableNode,
	TableCell,
} from "@core/table";
import { rejectShellMangledValue } from "../parse-helpers";
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

const AT_FORMS = describeForms(["table"], "                     ");

const HELP = `docx tables insert-row — insert a table row

Usage:
  docx tables insert-row FILE --at tN [options]

Required:
  --at LOCATOR       Target table. Supports:
${AT_FORMS}
                     See \`docx info locators\`.

Optional:
  --position INDEX   0-based row index to insert at (default: append at end)
  --cells "a,b,c"    Comma-separated text, one value per VISIBLE cell left to
                     right (default: empty). On a table with merged cells the new
                     row copies the neighbor row's gridSpan pattern, so you pass
                     one value per logical column (a spanned cell counts once),
                     NOT one per underlying grid column.
  --author NAME      Author for tracked insertion (default: $DOCX_AUTHOR)
  -o, --output PATH  Write to PATH instead of overwriting FILE
  --dry-run          Print what would change; do not write
  -v, --verbose      Print the success ack JSON
  -h, --help         Show this help

A row inserted inside a vertical merge extends the merge through the new row
(its cell in the merged column becomes a vMerge continuation) — matching Word.
A row inserted below the merge is a normal independent row.

When track-changes is on, the new row is wrapped as a tracked insertion
(<w:trPr><w:ins/>) — accept keeps it, reject removes it.

Output:
  Prints a one-line confirmation on success (exit 0). --verbose prints {ok:true, operation, path, table,
  position, tracked}. --dry-run prints the preview object (no ok field). Errors
  print {code, error, hint?} with a nonzero exit.

Examples:
  docx tables insert-row doc.docx --at t0
  docx tables insert-row doc.docx --at t0 --position 1 --cells "Q3,42,up"
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			position: { type: "string" },
			cells: { type: "string" },
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
	if (grid.colCount < 1) {
		return fail("TABLE_STRUCTURE", `Table ${tableId} has no columns`);
	}

	const position = resolveRowPosition(parsed.values.position, grid.rows.length);
	if (position === null) {
		return fail(
			"USAGE",
			`--position must be an integer in 0..${grid.rows.length}`,
		);
	}

	const cellTexts = parsed.values.cells
		? (parsed.values.cells as string).split(",")
		: [];
	for (const cell of cellTexts) {
		const mangled = await rejectShellMangledValue(cell, HELP, "--cells");
		if (typeof mangled === "number") return mangled;
	}
	// Validate against the row's LOGICAL columns (a merged/gridSpan cell counts
	// once), not the raw grid-column count — the new row mirrors the sibling row's
	// span pattern, so `--cells` maps one value per visible cell.
	const reference = referenceRow(grid, position);
	const logicalCols = reference ? reference.cells.length : grid.colCount;
	if (cellTexts.length > logicalCols) {
		return fail(
			"USAGE",
			`--cells has ${cellTexts.length} entries but the row has ${logicalCols} columns (a merged cell counts once)`,
		);
	}

	const newRow = buildRow(grid, position, cellTexts);

	const tracking = resolveTracked(document, parsed.values.track);
	if (tracking) {
		const meta = new TrackChanges(document).mintMeta(
			parsed.values.author as string | undefined,
		);
		markRowTracked(newRow, "ins", meta);
	}

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			operation: "tables.insert-row",
			dryRun: true,
			path,
			table: tableId,
			position,
			tracked: tracking,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	const childIndex = rowChildIndex(tableNode, grid, position);
	tableNode.children.splice(childIndex, 0, newRow);

	await document.save(outputPath);

	await respondAck({
		ok: true,
		operation: "tables.insert-row",
		path: outputPath ?? path,
		table: tableId,
		position,
		tracked: tracking,
	});
	return EXIT.OK;
}

function resolveRowPosition(raw: unknown, rowCount: number): number | null {
	if (raw === undefined) return rowCount;
	const value = Number(raw);
	if (!Number.isInteger(value) || value < 0 || value > rowCount) return null;
	return value;
}

/** The sibling row whose horizontal `gridSpan` pattern the new row copies — the
 * row ABOVE the insertion point (a data row, the canonical structure), falling
 * back to the row below when inserting at the top, or undefined for an empty
 * table. Also drives the `--cells` logical-column count. */
function referenceRow(grid: Grid, position: number): GridRow | undefined {
	const above = position > 0 ? grid.rows[position - 1] : undefined;
	const below = position < grid.rows.length ? grid.rows[position] : undefined;
	return above ?? below;
}

/** Build the `<w:tr>` to insert at logical `position`, mirroring the table's
 * existing column structure rather than emitting a flat band of single cells.
 *
 * Two patterns are honored so the new row lines up with its neighbors:
 *  - **Horizontal `gridSpan`** is copied from the reference sibling row, so a
 *    table whose data rows merge two grid columns (e.g. an invoice "Quantity"
 *    cell spanning 2) gets a matching spanned cell — NOT two stray cells that
 *    shove every later value one column left.
 *  - **Vertical merge** spanning the insertion point still gets a
 *    `vMerge="continue"` cell so the merge extends through the new row (Word's
 *    behavior); a continuation isn't an editable cell, so it consumes no
 *    `--cells` value.
 *
 * `--cells` maps one value per LOGICAL (visible) cell, left to right. By
 * construction the emitted spans sum to `grid.colCount`, so the row always
 * matches `<w:tblGrid>`. */
function buildRow(grid: Grid, position: number, cellTexts: string[]): XmlNode {
	const below =
		position > 0 && position < grid.rows.length
			? grid.rows[position]
			: undefined;
	const spans = referenceRow(grid, position);
	const cells: XmlNode[] = [];
	let logical = 0;
	for (let col = 0; col < grid.colCount; ) {
		const continued = below ? cellAt(below, col) : undefined;
		if (continued?.vMerge === "continue") {
			cells.push(
				<TableCell
					vMerge="continue"
					gridSpan={continued.colSpan > 1 ? continued.colSpan : undefined}
				>
					<Paragraph text="" />
				</TableCell>,
			);
			col += continued.colSpan;
			continue;
		}
		const refSpan = spans ? (cellAt(spans, col)?.colSpan ?? 1) : 1;
		cells.push(
			<TableCell gridSpan={refSpan > 1 ? refSpan : undefined}>
				<Paragraph text={cellTexts[logical] ?? ""} />
			</TableCell>,
		);
		logical += 1;
		col += refSpan;
	}
	return <w.tr>{cells}</w.tr>;
}

/** Index in `table.children` at which to splice a row inserted at logical
 * `position`, skipping past tblPr/tblGrid and any preceding rows. */
function rowChildIndex(
	table: XmlNode,
	grid: ReturnType<typeof buildGrid>,
	position: number,
): number {
	if (position < grid.rows.length) {
		return table.children.indexOf(grid.rows[position]?.node as XmlNode);
	}
	const lastRow = grid.rows[grid.rows.length - 1]?.node;
	return lastRow ? table.children.indexOf(lastRow) + 1 : table.children.length;
}
