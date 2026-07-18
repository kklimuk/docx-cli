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
	RENDER_VERIFY_EXAMPLE,
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

Examples:
  docx tables set-widths doc.docx --at t0 --widths "25%,25%,50%"
  docx tables set-widths doc.docx --at t0 --widths "1440,1440,2880"
  docx tables set-widths doc.docx --at t0 --widths auto
  # AGENT VERIFICATION: widths don't show in \`docx read\` — look at the pages
${RENDER_VERIFY_EXAMPLE}

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

Percentages and twips set a fixed layout. Under track-changes the resize is
recorded as a real revision, so Word can accept or reject it.

Widths map one value per GRID column, not per visible column. On a table with
merged cells, the grid has MORE columns than a single row shows, so the count
you pass must match the grid (run \`docx read --ast\` to see it). A cell that
HOLDS TEXT but lands narrower than ~0.2in is refused — that's under one
character wide, so Word would wrap it one char per line (empty/spacer columns
that thin are fine). A wider cell whose longest value still won't fit (e.g. a
0.78in column holding "$10,100.00") is applied but prints a WARNING naming the
column and the value — heed it: the value wraps mid-number in Word.

Output:
  Prints a one-line confirmation echoing the resulting per-column widths
  (exit 0). --verbose prints {ok:true, operation, path, table, layout,
  widths}. --dry-run prints the preview object (no ok field). Errors print
  {code, error, hint?} with a nonzero exit.
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
	let wrapWarnings: string[] = [];
	if (!auto) {
		const resolved = resolveWidths(widthsSpec, cols, grid);
		if (typeof resolved === "string") return fail("USAGE", resolved);
		twips = resolved;
		// One walk over the table's text-bearing cells feeds BOTH checks, computed
		// here (before the dry-run branch) so `--dry-run` previews the same
		// warnings a real run prints, and so each cell's text is collected once,
		// not twice. Refusal: a cell too narrow to fit one character (Word wraps
		// it one char per line — a render-only break `read` won't show). Warning:
		// a cell whose longest token still overflows the assigned width.
		const textCells = [...textBearingCells(grid, twips)];
		const tooNarrow = findTooNarrowCell(textCells);
		if (tooNarrow) return fail("USAGE", tooNarrow);
		wrapWarnings = findWrapRiskCells(
			textCells,
			baselineSizeHalfPoints(document),
		);
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
			...(wrapWarnings.length ? { warnings: wrapWarnings } : {}),
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
	const warningText = wrapWarnings.length
		? `${wrapWarnings.map((warning) => `WARNING: ${warning}`).join("\n")}\n`
		: "";
	const echo = auto ? "" : `${describeColumnWidths(twips)}\n`;
	await respondAck(
		{
			ok: true,
			operation: "tables.set-widths",
			path: destination,
			table: tableId,
			layout: auto ? "autofit" : "fixed",
			widths: auto ? "auto" : twips,
			...(wrapWarnings.length ? { warnings: wrapWarnings } : {}),
		},
		`${echo}${warningText}${renderVerifyHint(destination)}`,
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

function findTooNarrowCell(cells: readonly TextBearingCell[]): string | null {
	for (const { cell, widthTwips, text } of cells) {
		if (widthTwips >= MIN_COL_TWIPS) continue;
		const where =
			cell.colSpan > 1
				? `grid columns ${cell.colStart}–${cell.colStart + cell.colSpan - 1}`
				: `grid column ${cell.colStart}`;
		return `--widths collapses ${where} to ${twipsToInches(widthTwips)}in (${widthTwips} twips); that cell holds "${truncateSample(text)}" but ~0.15in goes to cell margins, leaving under one character — Word wraps it one char per line. Widen it and lower a wider column to compensate.`;
	}
	return null;
}

type TextBearingCell = { cell: GridCell; widthTwips: number; text: string };

/** One walk over the grid's text-bearing cells: each with its resolved width
 * (merge-aware — a spanning cell sums its grid columns) and trimmed text.
 * Materialized once and shared by the refusal (`findTooNarrowCell`) and warning
 * (`findWrapRiskCells`) passes so the merge/empty-spacer knowledge lives once
 * and each cell's text is collected once, not per pass. */
function* textBearingCells(
	grid: Grid,
	twips: number[],
): Generator<TextBearingCell> {
	for (const row of grid.rows) {
		for (const cell of row.cells) {
			const text = cell.node.collectText().trim();
			if (text.length === 0) continue; // empty/spacer cell never wraps
			yield { cell, widthTwips: cellWidth(cell, twips), text };
		}
	}
}

/** The document's docDefaults font size in half-points, guarded against a
 * malformed `w:sz` (which `defaultSizeHalfPoints` returns as NaN): fall back to
 * the 11pt (22 half-point) template baseline. A NaN would slip through `?? 22`
 * and poison the wrap-risk estimate (`neededTwips` NaN → every column flagged,
 * "~NaNin" in the ack). */
function baselineSizeHalfPoints(document: Document): number {
	const size = document.styles?.defaultSizeHalfPoints();
	return typeof size === "number" && Number.isFinite(size) && size > 0
		? size
		: 22;
}

function truncateSample(text: string): string {
	return text.length > 24 ? `${text.slice(0, 24)}…` : text;
}

/** True if any cell spans more than one grid column — the table where "visible
 * columns" and "grid columns" diverge and `--widths` becomes a footgun. */
function hasMergedColumns(grid: Grid): boolean {
	return grid.rows.some((row) => row.cells.some((cell) => cell.colSpan > 1));
}

/** Predict cells whose longest unbreakable token won't fit the new width, so
 * the wrap is flagged AT MUTATION TIME instead of surfacing only in a Word
 * render nobody runs (the invoice batch defect: a 0.78in Amount column
 * "accepted" $10,100.00 and wrapped every dollar value mid-number). Word wraps
 * at spaces; a space-free token wider than the usable cell width breaks
 * mid-token. Width is estimated at ~0.5em per character (the proportional-face
 * average; digits in Calibri are ≈0.507em) from the cell's own run size, so
 * this is a heuristic — hence a WARNING on the ack, not a refusal: the widths
 * the caller named are still applied. One warning per grid column (worst cell
 * wins), capped at 3 so the ack stays readable. */
function findWrapRiskCells(
	cells: readonly TextBearingCell[],
	baselineSizeHalfPoints: number,
): string[] {
	const CELL_MARGIN_TWIPS = 216; // 108/side Word default
	const worstByColumn = new Map<
		number,
		{ token: string; neededTwips: number; widthTwips: number }
	>();
	for (const { cell, widthTwips, text } of cells) {
		if (widthTwips < MIN_COL_TWIPS) continue; // already refused above
		const longest = text
			.split(/\s+/)
			.reduce((best, token) => (token.length > best.length ? token : best), "");
		if (!longest) continue;
		// Average char width ≈ 0.5em = (sz/2 half-points → pt)/2 = sz/4 pt;
		// in twips (×20): sz × 5 per character.
		const sizeHalfPoints = cellRunSize(cell.node, baselineSizeHalfPoints);
		const neededTwips = longest.length * sizeHalfPoints * 5 + CELL_MARGIN_TWIPS;
		if (neededTwips <= widthTwips) continue;
		const existing = worstByColumn.get(cell.colStart);
		if (!existing || neededTwips > existing.neededTwips) {
			worstByColumn.set(cell.colStart, {
				token: longest,
				neededTwips,
				widthTwips,
			});
		}
	}
	return [...worstByColumn.entries()]
		.sort(([, a], [, b]) => b.neededTwips - a.neededTwips)
		.slice(0, 3)
		.map(
			([colStart, { token, neededTwips, widthTwips }]) =>
				`grid column ${colStart} is ${twipsToInches(widthTwips)}in but holds "${truncateSample(token)}" (~${twipsToInches(neededTwips)}in incl. margins) — it will wrap mid-value in Word. Widen it and shrink a roomier column.`,
		);
}

/** The cell's first run's font size in half-points (Word's `w:sz` unit),
 * falling back to the document's docDefaults size when no run states one. */
function cellRunSize(cell: XmlNode, baselineSizeHalfPoints: number): number {
	const size = cell
		.findChild("w:p")
		?.findChild("w:r")
		?.findChild("w:rPr")
		?.findChild("w:sz")
		?.getAttribute("w:val");
	const parsed = Number(size);
	return Number.isFinite(parsed) && parsed > 0
		? parsed
		: baselineSizeHalfPoints;
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
