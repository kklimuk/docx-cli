import { describeForms, TrackChanges } from "@core";
import { parseRowAt } from "@core/locators";
import {
	buildGrid,
	cellAt,
	markRowTracked,
	resolveTableNode,
} from "@core/table";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const AT_FORMS = describeForms(["tableRow"], "                     ");

const HELP = `docx tables delete-row — delete a table row

Usage:
  docx tables delete-row FILE --at tN:rR [options]

Required:
  --at LOCATOR       Row to delete. Supports:
${AT_FORMS}
                     See \`docx info locators\`.

Optional:
  --author NAME      Author for tracked deletion (default: $DOCX_AUTHOR)
  -o, --output PATH  Write to PATH instead of overwriting FILE
  --dry-run          Print what would change; do not write
  -v, --verbose      Print the success ack JSON
  -h, --help         Show this help

When track-changes is on, the row is marked as a tracked deletion
(<w:trPr><w:del/>) rather than removed — accept removes it, reject keeps it.

Rejected if the row holds the "restart" half of a vertical merge whose
continuation rows would be orphaned; unmerge first.

Output:
  Silent on success (exit 0). --verbose prints {ok:true, operation, path, table,
  row, tracked}. --dry-run prints the preview object (no ok field). Errors print
  {code, error, hint?} with a nonzero exit.

Examples:
  docx tables delete-row doc.docx --at t0:r2
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			author: { type: "string" },
			output: { type: "string", short: "o" },
			"dry-run": { type: "boolean" },
			verbose: { type: "boolean", short: "v" },
			help: { type: "boolean", short: "h" },
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
	if (!at) return fail("USAGE", "Missing --at tN:rR", HELP);
	const target = parseRowAt(at);
	if (!target) {
		return fail(
			"INVALID_LOCATOR",
			`--at must be a row locator like t0:r1 (got ${at})`,
		);
	}

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const tableNode = resolveTableNode(document, target.tableId);
	if (!tableNode) {
		return fail("BLOCK_NOT_FOUND", `Table not found: ${target.tableId}`);
	}

	const grid = buildGrid(tableNode);
	const row = grid.rows[target.row];
	if (!row) {
		return fail(
			"BLOCK_NOT_FOUND",
			`Row ${target.row} not found in ${target.tableId} (has ${grid.rows.length} rows)`,
		);
	}

	// Guard both modes: untracked deletion orphans immediately, and a tracked
	// deletion orphans on accept (we don't promote the next continuation to
	// restart, unlike Word) — so refuse either way and point at unmerge.
	const orphanColumn = findVMergeOrphan(grid, target.row);
	if (orphanColumn !== null) {
		return fail(
			"TABLE_STRUCTURE",
			`Deleting row ${target.row} would orphan a vertical merge at column ${orphanColumn}`,
			`Run \`docx tables unmerge ${target.tableId}:r${target.row}c${orphanColumn}\` first.`,
		);
	}

	const tracking = document.isTrackChangesEnabled();
	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			operation: "tables.delete-row",
			dryRun: true,
			path,
			table: target.tableId,
			row: target.row,
			tracked: tracking,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	if (tracking) {
		markRowTracked(
			row.node,
			"del",
			new TrackChanges(document).mintMeta(
				parsed.values.author as string | undefined,
			),
		);
	} else {
		const index = tableNode.children.indexOf(row.node);
		if (index !== -1) tableNode.children.splice(index, 1);
	}

	await document.save(outputPath);

	await respondAck({
		ok: true,
		operation: "tables.delete-row",
		path: outputPath ?? path,
		table: target.tableId,
		row: target.row,
		tracked: tracking,
	});
	return EXIT.OK;
}

/** Return the logical column whose vertical merge would be orphaned by deleting
 * `rowIndex` — a "restart" cell here with a "continue" directly below — or null
 * when the deletion is safe. */
function findVMergeOrphan(
	grid: ReturnType<typeof buildGrid>,
	rowIndex: number,
): number | null {
	const row = grid.rows[rowIndex];
	const below = grid.rows[rowIndex + 1];
	if (!row || !below) return null;
	for (const cell of row.cells) {
		if (cell.vMerge !== "restart") continue;
		const continuation = cellAt(below, cell.colStart);
		if (continuation?.vMerge === "continue") return cell.colStart;
	}
	return null;
}
