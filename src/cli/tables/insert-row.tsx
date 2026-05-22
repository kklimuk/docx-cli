import { isTrackChangesEnabled, saveDocView } from "@core";
import { Paragraph } from "@core/blocks";
import { w } from "@core/jsx";
import { LocatorParseError, parseLocator } from "@core/locators";
import type { XmlNode } from "@core/parser";
import { TableCell } from "@core/table";
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
import { mintRevisionMeta, resolveTableNode } from "./common";
import { buildGrid, cellAt, type Grid } from "./grid";
import { markRowTracked } from "./mutate";

const HELP = `docx tables insert-row — insert a table row

Usage:
  docx tables insert-row FILE --at tN [options]

Required:
  --at tN            Target table (e.g. t0)

Optional:
  --position INDEX   0-based row index to insert at (default: append at end)
  --cells "a,b,c"    Comma-separated text for the new cells (default: empty).
                     Must not exceed the table's column count.
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

Examples:
  docx tables insert-row doc.docx --at t0
  docx tables insert-row doc.docx --at t0 --position 1 --cells "Q3,42,up"
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
				cells: { type: "string" },
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
	const tableId = tableIdFromArg(at);
	if (!tableId) {
		return fail(
			"INVALID_LOCATOR",
			`--at must be a table id like t0 (got ${at})`,
		);
	}

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const tableNode = resolveTableNode(view, tableId);
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
	if (cellTexts.length > grid.colCount) {
		return fail(
			"USAGE",
			`--cells has ${cellTexts.length} entries but table ${tableId} has ${grid.colCount} columns`,
		);
	}

	const newRow = buildRow(grid, position, cellTexts);

	const tracking = isTrackChangesEnabled(view);
	if (tracking) {
		const meta = mintRevisionMeta(
			view,
			parsed.values.author as string | undefined,
		);
		markRowTracked(newRow, "ins", meta);
	}

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
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

	await saveDocView(view, outputPath);

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

function tableIdFromArg(at: string): string | null {
	try {
		const locator = parseLocator(at);
		if (locator.kind === "block" && /^t\d+$/.test(locator.blockId)) {
			return locator.blockId;
		}
	} catch (error) {
		if (!(error instanceof LocatorParseError)) throw error;
	}
	return null;
}

function resolveRowPosition(raw: unknown, rowCount: number): number | null {
	if (raw === undefined) return rowCount;
	const value = Number(raw);
	if (!Number.isInteger(value) || value < 0 || value > rowCount) return null;
	return value;
}

/** Build the `<w:tr>` to insert at logical `position`. Columns whose vertical
 * merge spans the insertion point get a `vMerge="continue"` cell so the merge
 * *extends* through the new row (matching Word's behavior — inserting inside a
 * vertical merge grows it rather than splitting it); other columns get a normal
 * cell carrying the matching `--cells` text. A merge spans the insertion point
 * when the row now below (`grid.rows[position]`) is a `continue` there — which
 * never happens at the top (0) or at append (rows.length). */
function buildRow(grid: Grid, position: number, cellTexts: string[]): XmlNode {
	const below =
		position > 0 && position < grid.rows.length
			? grid.rows[position]
			: undefined;
	const cells: XmlNode[] = [];
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
		cells.push(
			<TableCell>
				<Paragraph text={cellTexts[col] ?? ""} />
			</TableCell>,
		);
		col += 1;
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
