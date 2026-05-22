import { saveDocView } from "@core";
import { LocatorParseError, parseLocator } from "@core/locators";
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
import { noteStructuralChange, resolveTableNode } from "./common";
import { buildGrid, cellAt, type Grid } from "./grid";
import { emptyCell, setGridSpan, setVMerge } from "./mutate";

const HELP = `docx tables unmerge — split a merged cell back into individual cells

Usage:
  docx tables unmerge FILE --at tN:rRcC [options]

Required:
  --at tN:rRcC       The merge anchor — the top-left cell of the merge
                     (the <w:gridSpan> cell and/or the "restart" of a vMerge)

Optional:
  --author NAME      Author for the audit comment when track-changes is on
  -o, --output PATH  Write to PATH instead of overwriting FILE
  --dry-run          Print what would change; do not write
  -v, --verbose      Print the success ack JSON
  -h, --help         Show this help

Horizontal spans are split by re-inserting the collapsed empty cells; vertical
merges are split by stripping the <w:vMerge> markers from the anchor and its
continuation cells. OOXML has no tracked-change construct for this, so under
track-changes the change applies immediately with a [docx-cli] audit comment.

Examples:
  docx tables unmerge doc.docx --at t0:r0c0
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
	if (!at) return fail("USAGE", "Missing --at tN:rRcC", HELP);
	const target = parseCellArg(at);
	if (!target) {
		return fail(
			"INVALID_LOCATOR",
			`--at must be a cell locator like t0:r0c0 (got ${at})`,
		);
	}

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const tableNode = resolveTableNode(view, target.tableId);
	if (!tableNode) {
		return fail("BLOCK_NOT_FOUND", `Table not found: ${target.tableId}`);
	}

	const grid = buildGrid(tableNode);
	const row = grid.rows[target.row];
	const anchor = row ? cellAt(row, target.col) : undefined;
	if (!row || !anchor) {
		return fail(
			"BLOCK_NOT_FOUND",
			`Cell ${at} not found (table is ${grid.rows.length}×${grid.colCount})`,
		);
	}
	if (anchor.vMerge === "continue") {
		return fail(
			"TABLE_STRUCTURE",
			`${at} is a vertical-merge continuation, not the anchor`,
			"Target the top (restart) cell of the merge.",
		);
	}
	const horizontal = anchor.colSpan > 1;
	const vertical = anchor.vMerge === "restart";
	if (!horizontal && !vertical) {
		return fail("TABLE_STRUCTURE", `${at} is not a merged cell`);
	}

	const spanRows = verticalSpanRows(grid, target.row, target.col, vertical);
	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "tables.unmerge",
			dryRun: true,
			path,
			table: target.tableId,
			cell: { row: target.row, col: target.col },
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	for (const rowIndex of spanRows) {
		const spanRow = grid.rows[rowIndex];
		const cell = spanRow ? cellAt(spanRow, target.col) : undefined;
		if (!spanRow || !cell) continue;
		setVMerge(cell.node, null);
		if (cell.colSpan > 1) {
			setGridSpan(cell.node, 1);
			const index = spanRow.node.children.indexOf(cell.node);
			const fillers = Array.from({ length: cell.colSpan - 1 }, () =>
				emptyCell(),
			);
			spanRow.node.children.splice(index + 1, 0, ...fillers);
		}
	}

	// Word applies unmerge immediately even under tracking (no revision marker),
	// so we do the same and note it with an audit comment.
	noteStructuralChange(
		view,
		anchor.node,
		`cell unmerged (r${target.row}c${target.col})`,
		parsed.values.author as string | undefined,
	);

	await saveDocView(view, outputPath);

	await respondAck({
		ok: true,
		operation: "tables.unmerge",
		path: outputPath ?? path,
		table: target.tableId,
		cell: { row: target.row, col: target.col },
	});
	return EXIT.OK;
}

function parseCellArg(
	at: string,
): { tableId: string; row: number; col: number } | null {
	try {
		const locator = parseLocator(at);
		if (locator.kind === "cell") {
			return { tableId: locator.tableId, row: locator.row, col: locator.col };
		}
	} catch (error) {
		if (!(error instanceof LocatorParseError)) throw error;
	}
	return null;
}

/** Rows covered by a vertical merge starting at (row, col): the anchor row plus
 * every following row whose cell at `col` is a "continue". Just the anchor row
 * when the merge is horizontal-only. */
function verticalSpanRows(
	grid: Grid,
	row: number,
	col: number,
	vertical: boolean,
): number[] {
	const rows = [row];
	if (!vertical) return rows;
	for (let index = row + 1; index < grid.rows.length; index++) {
		if (cellAt(grid.rows[index], col)?.vMerge === "continue") rows.push(index);
		else break;
	}
	return rows;
}
