import {
	type DocView,
	isTrackChangesEnabled,
	mintRevisionMeta,
	saveDocView,
} from "@core";
import { parseTableAt } from "@core/locators";
import type { XmlNode } from "@core/parser";
import {
	appendTblGridChange,
	appendTcPrChange,
	buildGrid,
	type Grid,
	type GridCell,
	resolveTableNode,
	setCellWidth,
	setTableLayout,
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

const HELP = `docx tables set-widths — set column widths

Usage:
  docx tables set-widths FILE --at tN --widths SPEC [options]

Required:
  --at tN            Target table (e.g. t0)
  --widths SPEC      One of:
                       "20%,30%,50%"  percentages (must sum to ~100)
                       "1440,2880"    per-column twips
                       "auto"         switch the table to autofit layout

Optional:
  --author NAME      Author for the tracked change (default: $DOCX_AUTHOR)
  -o, --output PATH  Write to PATH instead of overwriting FILE
  --dry-run          Print what would change; do not write
  -v, --verbose      Print the success ack JSON
  -h, --help         Show this help

Percentages and twips set a fixed layout and rewrite <w:tblGrid> plus each
cell's <w:tcW>. Under track-changes the resize is recorded as a real revision
(<w:tblGridChange> for the grid plus a per-cell <w:tcPrChange>), so it can be
accepted or rejected — matching what Word emits for a width change.

Examples:
  docx tables set-widths doc.docx --at t0 --widths "25%,25%,50%"
  docx tables set-widths doc.docx --at t0 --widths "1440,1440,2880"
  docx tables set-widths doc.docx --at t0 --widths auto
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				at: { type: "string" },
				widths: { type: "string" },
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

	const widthsSpec = parsed.values.widths as string | undefined;
	if (!widthsSpec) return fail("USAGE", "Missing --widths", HELP);

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const tableNode = resolveTableNode(view, tableId);
	if (!tableNode) return fail("BLOCK_NOT_FOUND", `Table not found: ${tableId}`);

	const grid = buildGrid(tableNode);
	if (!grid.tblGrid) {
		return fail("TABLE_STRUCTURE", `Table ${tableId} has no <w:tblGrid>`);
	}

	const auto = widthsSpec.trim() === "auto";
	const cols = grid.tblGrid.findChildren("w:gridCol");
	let twips: number[] = [];
	if (!auto) {
		const resolved = resolveWidths(widthsSpec, cols);
		if (typeof resolved === "string") return fail("USAGE", resolved);
		twips = resolved;
	}

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "tables.set-widths",
			dryRun: true,
			path,
			table: tableId,
			layout: auto ? "autofit" : "fixed",
			widths: auto ? "auto" : twips,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	const tracking = isTrackChangesEnabled(view);
	const author = parsed.values.author as string | undefined;
	// Snapshot the prior grid columns before resizing so a tracked width change
	// is reversible (reject restores the prior <w:tblGrid> from the snapshot).
	const priorCols = tracking ? cols.map((col) => col.clone()) : [];

	if (auto) {
		setTableLayout(tableNode, "autofit");
	} else {
		cols.forEach((col, index) => {
			const value = twips[index];
			if (value !== undefined) col.setAttribute("w:w", String(value));
		});
		// Word records a width change as a grid revision (<w:tblGridChange>) plus
		// a per-cell <w:tcPrChange> (each cell's <w:tcW>) — and it's the per-cell
		// tcPrChange that Word's reject actually reverts (a grid snapshot alone
		// isn't honored). Mirror Word's full output under tracking.
		if (tracking) applyCellWidthsTracked(view, grid, twips, author);
		else applyCellWidths(grid, twips);
		setTableLayout(tableNode, "fixed");
	}

	if (tracking && grid.tblGrid) {
		appendTblGridChange(
			grid.tblGrid,
			priorCols,
			mintRevisionMeta(view, author),
		);
	}

	await saveDocView(view, outputPath);

	await respondAck({
		ok: true,
		operation: "tables.set-widths",
		path: outputPath ?? path,
		table: tableId,
		layout: auto ? "autofit" : "fixed",
		widths: auto ? "auto" : twips,
	});
	return EXIT.OK;
}

/** Resolve a `--widths` spec to per-column twips, or an error message. */
function resolveWidths(spec: string, cols: XmlNode[]): number[] | string {
	const tokens = spec.split(",").map((token) => token.trim());
	if (tokens.length !== cols.length) {
		return `--widths has ${tokens.length} entries but the table has ${cols.length} columns`;
	}
	const percentage = tokens.every((token) => token.endsWith("%"));
	const anyPercent = tokens.some((token) => token.endsWith("%"));
	if (anyPercent && !percentage) {
		return "--widths must be all percentages or all twips, not mixed";
	}
	if (percentage) {
		const percents = tokens.map((token) => Number(token.slice(0, -1)));
		if (percents.some((value) => !Number.isFinite(value) || value <= 0)) {
			return "--widths percentages must be positive numbers";
		}
		const sum = percents.reduce((total, value) => total + value, 0);
		if (Math.abs(sum - 100) > 1) {
			return `--widths percentages must sum to ~100 (got ${sum})`;
		}
		const total = currentTotal(cols);
		return percents.map((value) => Math.round((value / 100) * total));
	}
	const values = tokens.map((token) => Number(token));
	if (values.some((value) => !Number.isInteger(value) || value <= 0)) {
		return "--widths twips must be positive integers";
	}
	return values;
}

function currentTotal(cols: XmlNode[]): number {
	const sum = cols.reduce(
		(total, col) => total + Number(col.getAttribute("w:w") ?? "0"),
		0,
	);
	return sum > 0 ? sum : 9360;
}

/** Rewrite each cell's <w:tcW> to the sum of the grid widths it spans, so cell
 * widths stay consistent with the grid (merged cells get the combined width). */
function applyCellWidths(grid: Grid, twips: number[]): void {
	for (const row of grid.rows) {
		for (const cell of row.cells) {
			setCellWidth(cell.node, { value: cellWidth(cell, twips), unit: "dxa" });
		}
	}
}

/** As {@link applyCellWidths}, but records a `<w:tcPrChange>` snapshot of each
 * cell's prior `<w:tcPr>` — this is the revision Word's reject reverts. */
function applyCellWidthsTracked(
	view: DocView,
	grid: Grid,
	twips: number[],
	authorFlag: string | undefined,
): void {
	for (const row of grid.rows) {
		for (const cell of row.cells) {
			// Order matters: snapshot the prior tcPr BEFORE mutating, then append
			// the change AFTER — reordering would snapshot the post-change state.
			const prior = (cell.node.findChild("w:tcPr")?.children ?? []).map(
				(child) => child.clone(),
			);
			setCellWidth(cell.node, { value: cellWidth(cell, twips), unit: "dxa" });
			appendTcPrChange(cell.node, prior, mintRevisionMeta(view, authorFlag));
		}
	}
}

function cellWidth(cell: GridCell, twips: number[]): number {
	let width = 0;
	for (let offset = 0; offset < cell.colSpan; offset++) {
		width += twips[cell.colStart + offset] ?? 0;
	}
	return width;
}
