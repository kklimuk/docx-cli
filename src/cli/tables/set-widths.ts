import { type Document, describeForms, TrackChanges } from "@core";
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
import { twipsToInches } from "../read/annotations";
import {
	EXIT,
	fail,
	openOrFail,
	renderVerifyHint,
	resolveTracked,
	respond,
	respondAck,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const AT_FORMS = describeForms(["table"], "                     ");

const HELP = `docx tables set-widths — set column widths

Usage:
  docx tables set-widths FILE --at tN --widths SPEC [options]

Required:
  --at LOCATOR       Target table. Supports:
${AT_FORMS}
                     See \`docx info locators\`.
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

Widths map one value per GRID column, not per visible column. On a table with
merged cells (gridSpan), the grid has MORE columns than a single row shows, so
the count you pass must match the grid (run \`docx read --ast\` to see it). A
cell that HOLDS TEXT but lands narrower than ~0.2in is refused — after ~0.15in
of cell margin it fits under one character, so Word wraps it one char per line
(empty/spacer columns that thin are fine, nothing to wrap). The success line
echoes the resulting per-column widths; since layout changes don't show in
\`read\`, render to verify.

Output:
  Prints a one-line confirmation on success (exit 0). --verbose prints {ok:true, operation, path, table,
  layout, widths}. --dry-run prints the preview object (no ok field). Errors
  print {code, error, hint?} with a nonzero exit.

Examples:
  docx tables set-widths doc.docx --at t0 --widths "25%,25%,50%"
  docx tables set-widths doc.docx --at t0 --widths "1440,1440,2880"
  docx tables set-widths doc.docx --at t0 --widths auto
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			widths: { type: "string" },
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

	const widthsSpec = parsed.values.widths as string | undefined;
	if (!widthsSpec) return fail("USAGE", "Missing --widths", HELP);

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const tableNode = resolveTableNode(document, tableId);
	if (!tableNode) return fail("BLOCK_NOT_FOUND", `Table not found: ${tableId}`);

	const grid = buildGrid(tableNode);
	if (!grid.tblGrid) {
		return fail("TABLE_STRUCTURE", `Table ${tableId} has no <w:tblGrid>`);
	}

	const auto = widthsSpec.trim() === "auto";
	const cols = grid.tblGrid.findChildren("w:gridCol");
	let twips: number[] = [];
	if (!auto) {
		const resolved = resolveWidths(widthsSpec, cols, grid);
		if (typeof resolved === "string") return fail("USAGE", resolved);
		twips = resolved;
		// Guard against a width that collapses a cell so narrow Word wraps its
		// content one character per line — a render-only break `read` won't show.
		const tooNarrow = findTooNarrowCell(grid, twips);
		if (tooNarrow) return fail("USAGE", tooNarrow);
	}

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
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

	const tracking = resolveTracked(document, parsed.values.track);
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
		if (tracking) applyCellWidthsTracked(document, grid, twips, author);
		else applyCellWidths(grid, twips);
		setTableLayout(tableNode, "fixed");
	}

	if (tracking && grid.tblGrid) {
		appendTblGridChange(
			grid.tblGrid,
			priorCols,
			new TrackChanges(document).mintMeta(author),
		);
	}

	await document.save(outputPath);

	const destination = outputPath ?? path;
	const echo = auto ? "" : `${describeColumnWidths(twips)}\n`;
	await respondAck(
		{
			ok: true,
			operation: "tables.set-widths",
			path: destination,
			table: tableId,
			layout: auto ? "autofit" : "fixed",
			widths: auto ? "auto" : twips,
		},
		`${echo}${renderVerifyHint(destination)}`,
	);
	return EXIT.OK;
}

/** Resolve a `--widths` spec to per-column twips, or an error message. Widths
 * map one value per GRID column (`<w:gridCol>`), not per VISIBLE column — on a
 * merged-cell table the two differ, and a positional list silently misaligns
 * (the invoice scenario's blocker). When the count mismatches a merged table we
 * explain the grid-vs-logical gap instead of the bare count. */
function resolveWidths(
	spec: string,
	cols: XmlNode[],
	grid: Grid,
): number[] | string {
	const tokens = spec.split(",").map((token) => token.trim());
	if (tokens.length !== cols.length) {
		const base = `--widths has ${tokens.length} entries but the table has ${cols.length} grid columns`;
		return hasMergedColumns(grid)
			? `${base}. This table has merged cells, so some visible columns span multiple grid columns and the grid has more columns than any single row shows. Supply one width per GRID column (${cols.length} values), left-to-right; a merged cell takes the sum of the grid columns it covers. Inspect the gridSpan layout with \`docx read --ast\`.`
			: base;
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
	document: Document,
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
			appendTcPrChange(
				cell.node,
				prior,
				new TrackChanges(document).mintMeta(authorFlag),
			);
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

/** ~0.2in (288 twips) isn't an arbitrary round number: Word eats ~0.15in of
 * default cell margin (108 twips/side) before any glyph, so a 0.2in cell leaves
 * under one character of usable width and wraps content one char per line — a
 * render-only break `read` never surfaces (the invoice scenario shipped a
 * 0.156in Amount cell exactly this way). The check is the actual rendering unit,
 * the CELL: a merged cell sums its grid columns (so a wide span isn't flagged),
 * and an EMPTY cell is skipped (a deliberate thin spacer column renders fine —
 * there's nothing to wrap). So we only refuse a narrow cell that holds text. */
const MIN_COL_TWIPS = 288;

function findTooNarrowCell(grid: Grid, twips: number[]): string | null {
	for (const row of grid.rows) {
		for (const cell of row.cells) {
			const width = cellWidth(cell, twips);
			if (width >= MIN_COL_TWIPS) continue;
			const text = cell.node.collectText().trim();
			if (text.length === 0) continue; // empty/spacer cell never wraps
			const where =
				cell.colSpan > 1
					? `grid columns ${cell.colStart}–${cell.colStart + cell.colSpan - 1}`
					: `grid column ${cell.colStart}`;
			const sample = text.length > 24 ? `${text.slice(0, 24)}…` : text;
			return `--widths collapses ${where} to ${twipsToInches(width)}in (${width} twips); that cell holds "${sample}" but ~0.15in goes to cell margins, leaving under one character — Word wraps it one char per line. Widen it and lower a wider column to compensate.`;
		}
	}
	return null;
}

/** True if any cell spans more than one grid column — the table where "visible
 * columns" and "grid columns" diverge and `--widths` becomes a footgun. */
function hasMergedColumns(grid: Grid): boolean {
	return grid.rows.some((row) => row.cells.some((cell) => cell.colSpan > 1));
}

/** Echo the resulting per-grid-column widths in inches so the agent can sanity-
 * check the assignment at the moment of success — the misaligned slot is then
 * self-evident (`g4=0.16in`) instead of invisible until a render. */
function describeColumnWidths(twips: number[]): string {
	const cells = twips.map(
		(value, index) => `g${index}=${twipsToInches(value)}in`,
	);
	return `widths: ${cells.join(" ")}`;
}
