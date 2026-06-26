import { type Document, Edit, TrackChanges } from "@core";
import type { ParagraphOptions } from "@core/blocks";
import {
	parseCellAt,
	parseCellRangeAt,
	parseColumnAt,
	parseRowAt,
	parseTableAt,
} from "@core/locators";
import type { XmlNode } from "@core/parser";
import {
	appendTcPrChange,
	type BorderEdge,
	buildGrid,
	type CellBorderSide,
	cellAt,
	type Grid,
	type GridCell,
	type GridRow,
	resolveTableNode,
	setCellBorders,
	setCellShading,
	setCellVAlign,
	setRepeatHeader,
	setRowHeight,
	setTableJustification,
	setTableStyle,
} from "@core/table";
import { normalizeHexColor } from "../parse-helpers";
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
import { noteStructuralChange } from "./common";

const HELP = `docx tables format — shade, align, border, and size table cells/rows/tables

Usage:
  docx tables format FILE --at LOCATOR [formatting options]

The --at locator picks WHAT to format; its granularity picks which options apply.
  --at t0            whole table
  --at t0:r2         a row          --at t0:c1   a column
  --at t0:r1c2       a single cell  --at t0:r1c0-r3c2   a rectangle of cells

Cell options (any locator — broadcast to every cell it covers):
  --shade HEX|NAME|none   Cell background fill (e.g. D9D9D9, "grey", or none to clear)
  --valign top|center|bottom   Vertical alignment of cell content
  --halign left|center|right|justify
                          Horizontal alignment of cell TEXT (a paragraph property —
                          not the same as --align, which moves the whole table)
  --cell-borders SIDES    Comma list of top,bottom,left,right,all,insideH,insideV
                          (or none to clear). Styled by the --border-* options:
  --border-style single|double|none   (default: single)
  --border-size EIGHTHS   Thickness in eighths of a point (default: 4 = 0.5pt)
  --border-color HEX|NAME (default: auto)

Row options (--at t0:rR, or --at t0 for every row):
  --row-height MEASURE    a unit is required: 0.4in or 28pt (a bare number is rejected)
  --height-rule atLeast|exact|auto   (default: atLeast)
  --repeat-header / --no-repeat-header
                          Repeat the row as a header atop each page the table spans

Table options (--at t0 only):
  --align left|center|right   Position the whole table on the page
  --style STYLEID|none    Apply (or clear) a table style from styles.xml

Common:
  --author NAME   Author for tracked changes / audit comments (default: $DOCX_AUTHOR)
  --track         Record as a tracked change even if the doc toggle is off
  -o, --output PATH / --dry-run / -v, --verbose / -h, --help

Tracking: cell shading/vAlign/borders record as a real <w:tcPrChange> and --halign
as a <w:pPrChange> (both round-trip in Word). Table align/style and row height/
repeat-header have no tracked-change construct Word will revert, so under tracking
they apply in place with a [docx-cli] audit comment (same policy as tables borders).

Examples:
  docx tables format invoice.docx --at t0:r0 --shade D9D9D9 --valign center
  docx tables format invoice.docx --at t0:c2 --halign right
  docx tables format invoice.docx --at t0:r1c0-r3c2 --cell-borders bottom
  docx tables format invoice.docx --at t0 --align center --style LightGrid
  docx tables format invoice.docx --at t0:r0 --repeat-header --row-height 0.4in
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			shade: { type: "string" },
			valign: { type: "string" },
			halign: { type: "string" },
			"cell-borders": { type: "string" },
			"border-style": { type: "string" },
			"border-size": { type: "string" },
			"border-color": { type: "string" },
			align: { type: "string" },
			style: { type: "string" },
			"row-height": { type: "string" },
			"height-rule": { type: "string" },
			"repeat-header": { type: "boolean" },
			"no-repeat-header": { type: "boolean" },
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
	if (!at) return fail("USAGE", "Missing --at LOCATOR", HELP);
	const scope = resolveScope(at);
	if (!scope) {
		return fail(
			"INVALID_LOCATOR",
			`--at must address a table, row, column, cell, or cell range (got ${at})`,
			"e.g. t0 (table), t0:r0 (row), t0:c1 (column), t0:r0c0 (cell), t0:r0c0-r1c2 (range).",
		);
	}

	const plan = buildPlan(parsed.values);
	if (typeof plan === "string") return fail("USAGE", plan);
	if (planIsEmpty(plan)) {
		return fail(
			"USAGE",
			"Nothing to format — pass at least one of --shade/--valign/--halign/--cell-borders/--align/--style/--row-height/--repeat-header",
		);
	}

	const scopeError = validateScope(scope, plan);
	if (scopeError) return fail("USAGE", scopeError);

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const tableNode = resolveTableNode(document, scope.tableId);
	if (!tableNode)
		return fail("BLOCK_NOT_FOUND", `Table not found: ${scope.tableId}`);
	const grid = buildGrid(tableNode);

	const cells = selectCells(grid, scope);
	const rows = selectRows(grid, scope);
	const targetError = validateTargets(scope, plan, cells, rows);
	if (targetError) return fail("BLOCK_NOT_FOUND", targetError);

	const outputPath = parsed.values.output as string | undefined;
	if (parsed.values["dry-run"]) {
		await respond({
			operation: "tables.format",
			dryRun: true,
			path,
			at,
			applied: appliedList(plan),
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	const tracking = resolveTracked(document, parsed.values.track);
	const author = parsed.values.author as string | undefined;

	applyCellProps(document, cells, plan, tracking, author);
	if (plan.halign !== undefined)
		applyHalign(document, cells, plan.halign, tracking, author);
	applyTableProps(tableNode, plan);
	applyRowProps(rows, plan);

	// Table-level (jc/style) and row-level (height/header) changes have no
	// tracked-change construct Word honors, so under tracking we mirror the
	// `borders` policy: apply in place + one [docx-cli] audit comment.
	const untracked = untrackedSummary(plan);
	if (untracked)
		noteStructuralChange(
			document,
			cells[0]?.node,
			untracked,
			author,
			parsed.values.track as boolean | undefined,
		);

	await document.save(outputPath);

	const destination = outputPath ?? path;
	await respondAck(
		{
			ok: true,
			operation: "tables.format",
			path: destination,
			table: at,
			applied: appliedList(plan),
		},
		renderVerifyHint(destination),
	);
	return EXIT.OK;
}

/** What `--at` addresses, resolved before the table is opened. Cell properties
 * broadcast over the covered cells; row/table properties need a row/table scope. */
type Scope =
	| { kind: "table"; tableId: string }
	| { kind: "row"; tableId: string; row: number }
	| { kind: "column"; tableId: string; col: number }
	| { kind: "cell"; tableId: string; row: number; col: number }
	| {
			kind: "range";
			tableId: string;
			start: { row: number; col: number };
			end: { row: number; col: number };
	  };

/** Try the locator parsers most-specific first. They're disjoint by locator
 * kind, so the first match is unambiguous. */
function resolveScope(at: string): Scope | null {
	const cell = parseCellAt(at);
	if (cell) return { kind: "cell", ...cell };
	const range = parseCellRangeAt(at);
	if (range) return { kind: "range", ...range };
	const row = parseRowAt(at);
	if (row) return { kind: "row", ...row };
	const column = parseColumnAt(at);
	if (column) return { kind: "column", ...column };
	const tableId = parseTableAt(at);
	if (tableId) return { kind: "table", tableId };
	return null;
}

/** The resolved formatting to apply. `undefined` = leave alone; `null` (on the
 * clearable props) = remove the property. */
type FormatPlan = {
	shade?: string | null;
	valign?: "top" | "center" | "bottom" | null;
	halign?: ParagraphOptions["alignment"];
	cellBorders?: {
		updates: { side: CellBorderSide; edge: BorderEdge | null }[];
		clearAll: boolean;
	};
	align?: "left" | "center" | "right" | null;
	style?: string | null;
	rowHeight?: { value: number; rule: "atLeast" | "exact" | "auto" } | null;
	repeatHeader?: boolean;
};

function buildPlan(values: Record<string, unknown>): FormatPlan | string {
	const plan: FormatPlan = {};

	const shade = values.shade as string | undefined;
	if (shade !== undefined) {
		if (isClear(shade)) plan.shade = null;
		else {
			const hex = resolveColor(shade);
			if (!hex) return colorError("--shade", shade);
			plan.shade = hex;
		}
	}

	const valign = values.valign as string | undefined;
	if (valign !== undefined) {
		if (!VALIGNS.has(valign)) return "--valign must be top, center, or bottom";
		plan.valign = valign as "top" | "center" | "bottom";
	}

	const halign = values.halign as string | undefined;
	if (halign !== undefined) {
		if (!HALIGNS.has(halign))
			return "--halign must be left, center, right, or justify";
		plan.halign = halign as ParagraphOptions["alignment"];
	}

	const cellBorders = values["cell-borders"] as string | undefined;
	if (cellBorders !== undefined) {
		const built = buildCellBorders(values, cellBorders);
		if (typeof built === "string") return built;
		plan.cellBorders = built;
	}

	const align = values.align as string | undefined;
	if (align !== undefined) {
		if (!TABLE_ALIGNS.has(align))
			return "--align must be left, center, or right";
		plan.align = align as "left" | "center" | "right";
	}

	const style = values.style as string | undefined;
	if (style !== undefined) plan.style = isClear(style) ? null : style;

	const rowHeight = values["row-height"] as string | undefined;
	const heightRule = values["height-rule"] as string | undefined;
	// Validate the rule independently so it can't be silently dropped when paired
	// with a non-height flag (e.g. `--repeat-header --height-rule exact`).
	if (heightRule !== undefined && !HEIGHT_RULES.has(heightRule))
		return "--height-rule must be atLeast, exact, or auto";
	if (heightRule !== undefined && rowHeight === undefined)
		return "--height-rule has no effect without --row-height";
	if (rowHeight !== undefined) {
		const twips = parseMeasureToTwips(rowHeight);
		if (twips === null)
			return `--row-height must be a positive measure WITH a unit, e.g. 0.4in or 28pt (got ${rowHeight})`;
		plan.rowHeight = {
			value: twips,
			rule: (heightRule ?? "atLeast") as "atLeast" | "exact" | "auto",
		};
	}

	if (values["repeat-header"] && values["no-repeat-header"])
		return "--repeat-header and --no-repeat-header are mutually exclusive";
	if (values["repeat-header"]) plan.repeatHeader = true;
	if (values["no-repeat-header"]) plan.repeatHeader = false;

	return plan;
}

function buildCellBorders(
	values: Record<string, unknown>,
	sidesRaw: string,
):
	| {
			updates: { side: CellBorderSide; edge: BorderEdge | null }[];
			clearAll: boolean;
	  }
	| string {
	if (isClear(sidesRaw)) return { updates: [], clearAll: true };
	const style = (values["border-style"] as string | undefined) ?? "single";
	if (!BORDER_STYLES.has(style))
		return "--border-style must be single, double, or none";
	const sizeRaw = values["border-size"] as string | undefined;
	if (
		sizeRaw !== undefined &&
		(!Number.isInteger(Number(sizeRaw)) || Number(sizeRaw) <= 0)
	)
		return "--border-size must be a positive integer (eighths of a point)";
	const sizeEighths = sizeRaw ? Number(sizeRaw) : 4;
	const colorRaw = (values["border-color"] as string | undefined) ?? "auto";
	const color = colorRaw === "auto" ? "auto" : resolveColor(colorRaw);
	if (!color) return colorError("--border-color", colorRaw);
	const edge: BorderEdge = { style, sizeEighths, color };

	const sides = new Set<CellBorderSide>();
	for (const token of sidesRaw.split(",").map((value) => value.trim())) {
		if (token === "all") {
			for (const side of ALL_SIDES) sides.add(side);
			continue;
		}
		if (!ALL_SIDES.includes(token as CellBorderSide))
			return `--cell-borders side "${token}" must be one of ${ALL_SIDES.join(", ")}, all, or none`;
		sides.add(token as CellBorderSide);
	}
	return {
		updates: [...sides].map((side) => ({ side, edge })),
		clearAll: false,
	};
}

function selectCells(grid: Grid, scope: Scope): GridCell[] {
	if (scope.kind === "table") return grid.rows.flatMap((row) => row.cells);
	if (scope.kind === "row") return grid.rows[scope.row]?.cells ?? [];
	if (scope.kind === "cell") {
		const cell = cellAt(grid.rows[scope.row], scope.col);
		return cell ? [cell] : [];
	}
	if (scope.kind === "column") {
		const out: GridCell[] = [];
		for (const row of grid.rows) {
			const cell = cellAt(row, scope.col);
			if (cell && !out.includes(cell)) out.push(cell);
		}
		return out;
	}
	// range — every physical cell whose span intersects the rectangle, deduped.
	const r1 = Math.min(scope.start.row, scope.end.row);
	const r2 = Math.max(scope.start.row, scope.end.row);
	const c1 = Math.min(scope.start.col, scope.end.col);
	const c2 = Math.max(scope.start.col, scope.end.col);
	const out: GridCell[] = [];
	for (let row = r1; row <= r2; row++) {
		for (const cell of grid.rows[row]?.cells ?? []) {
			if (cell.colStart <= c2 && cell.colStart + cell.colSpan - 1 >= c1)
				out.push(cell);
		}
	}
	return out;
}

function selectRows(grid: Grid, scope: Scope): GridRow[] {
	if (scope.kind === "table") return grid.rows;
	if (scope.kind === "row") {
		const row = grid.rows[scope.row];
		return row ? [row] : [];
	}
	return [];
}

/** Cell `<w:tcPr>` properties (shade/vAlign/borders) record natively: snapshot
 * the prior tcPr BEFORE mutating, apply, then one `<w:tcPrChange>` per cell —
 * mirroring `set-widths`'s tracked path (the revision Word's reject reverts). */
function applyCellProps(
	document: Document,
	cells: GridCell[],
	plan: FormatPlan,
	tracking: boolean,
	author: string | undefined,
): void {
	if (
		plan.shade === undefined &&
		plan.valign === undefined &&
		plan.cellBorders === undefined
	)
		return;
	for (const cell of cells) {
		const prior = tracking
			? (cell.node.findChild("w:tcPr")?.children ?? []).map((child) =>
					child.clone(),
				)
			: [];
		if (plan.shade !== undefined) setCellShading(cell.node, plan.shade);
		if (plan.valign !== undefined) setCellVAlign(cell.node, plan.valign);
		if (plan.cellBorders !== undefined)
			setCellBorders(
				cell.node,
				plan.cellBorders.updates,
				plan.cellBorders.clearAll,
			);
		if (tracking)
			appendTcPrChange(
				cell.node,
				prior,
				new TrackChanges(document).mintMeta(author),
			);
	}
}

/** Horizontal alignment is a PARAGRAPH property — broadcast `<w:jc>` to every
 * paragraph in the in-scope cells via the same `Edit.paragraphProperties` path
 * `edit --alignment` uses, so it records as a real `<w:pPrChange>` under tracking. */
function applyHalign(
	document: Document,
	cells: GridCell[],
	alignment: ParagraphOptions["alignment"],
	tracking: boolean,
	author: string | undefined,
): void {
	const edit = new Edit(document);
	for (const cell of cells) {
		for (const paragraph of cell.node.findChildren("w:p")) {
			edit.paragraphProperties(
				{ node: paragraph, parent: cell.node.children },
				{ alignment },
				{ authorFlag: author, track: tracking },
			);
		}
	}
}

function applyTableProps(tableNode: XmlNode, plan: FormatPlan): void {
	if (plan.align !== undefined) setTableJustification(tableNode, plan.align);
	if (plan.style !== undefined) setTableStyle(tableNode, plan.style);
}

function applyRowProps(rows: GridRow[], plan: FormatPlan): void {
	if (plan.rowHeight === undefined && plan.repeatHeader === undefined) return;
	for (const row of rows) {
		if (plan.rowHeight !== undefined) setRowHeight(row.node, plan.rowHeight);
		if (plan.repeatHeader !== undefined)
			setRepeatHeader(row.node, plan.repeatHeader);
	}
}

function validateScope(scope: Scope, plan: FormatPlan): string | null {
	const hasTableFlag = plan.align !== undefined || plan.style !== undefined;
	if (hasTableFlag && scope.kind !== "table")
		return "--align/--style position or restyle the whole table — use --at tN";
	const hasRowFlag =
		plan.rowHeight !== undefined || plan.repeatHeader !== undefined;
	if (hasRowFlag && scope.kind !== "table" && scope.kind !== "row")
		return "--row-height/--repeat-header apply to a row — use --at tN:rR (or --at tN for every row)";
	return null;
}

function validateTargets(
	scope: Scope,
	plan: FormatPlan,
	cells: GridCell[],
	rows: GridRow[],
): string | null {
	const hasCellFlag =
		plan.shade !== undefined ||
		plan.valign !== undefined ||
		plan.halign !== undefined ||
		plan.cellBorders !== undefined;
	if (hasCellFlag && cells.length === 0)
		return `No cells matched ${describeScope(scope)}`;
	const hasRowFlag =
		plan.rowHeight !== undefined || plan.repeatHeader !== undefined;
	if (hasRowFlag && rows.length === 0)
		return `No rows matched ${describeScope(scope)}`;
	return null;
}

function describeScope(scope: Scope): string {
	if (scope.kind === "table") return scope.tableId;
	if (scope.kind === "row") return `${scope.tableId}:r${scope.row}`;
	if (scope.kind === "column") return `${scope.tableId}:c${scope.col}`;
	if (scope.kind === "cell")
		return `${scope.tableId}:r${scope.row}c${scope.col}`;
	return `${scope.tableId}:r${scope.start.row}c${scope.start.col}-r${scope.end.row}c${scope.end.col}`;
}

/** Property names changed, for the success ack / dry-run preview. */
function appliedList(plan: FormatPlan): string[] {
	const out: string[] = [];
	if (plan.shade !== undefined) out.push("shade");
	if (plan.valign !== undefined) out.push("valign");
	if (plan.halign !== undefined) out.push("halign");
	if (plan.cellBorders !== undefined) out.push("cell-borders");
	if (plan.align !== undefined) out.push("align");
	if (plan.style !== undefined) out.push("style");
	if (plan.rowHeight !== undefined) out.push("row-height");
	if (plan.repeatHeader !== undefined) out.push("repeat-header");
	return out;
}

/** One-line description of the table/row changes that ride an audit comment
 * (everything except the natively-tracked cell props + halign), or "" if none. */
function untrackedSummary(plan: FormatPlan): string {
	const parts: string[] = [];
	if (plan.align !== undefined)
		parts.push(`table align ${plan.align ?? "left"}`);
	if (plan.style !== undefined)
		parts.push(
			plan.style ? `table style ${plan.style}` : "table style cleared",
		);
	if (plan.rowHeight !== undefined) parts.push("row height set");
	if (plan.repeatHeader !== undefined)
		parts.push(plan.repeatHeader ? "repeat-header on" : "repeat-header off");
	return parts.length > 0 ? `set ${parts.join("; ")}` : "";
}

function planIsEmpty(plan: FormatPlan): boolean {
	return appliedList(plan).length === 0;
}

function isClear(value: string): boolean {
	const lower = value.toLowerCase();
	return lower === "none" || lower === "auto";
}

/** Resolve a color word or hex to a 6-digit uppercase hex, or null when invalid.
 * A small name map keeps the common asks (`grey`, `yellow`) working for weak
 * agents that won't reach for a hex code. */
function resolveColor(value: string): string | null {
	const named = COLOR_NAMES[value.toLowerCase()];
	if (named) return named;
	const hex = normalizeHexColor(value).toUpperCase();
	return /^[0-9A-F]{6}$/.test(hex) ? hex : null;
}

function colorError(flag: string, value: string): string {
	return `${flag} must be a 6-digit hex (e.g. D9D9D9) or a name (${Object.keys(
		COLOR_NAMES,
	)
		.slice(0, 6)
		.join(", ")}, …) — got ${value}`;
}

/** Parse `0.4in` / `28pt` to twips. The unit is REQUIRED — a bare number is
 * rejected rather than guessed, so `--row-height 28` can't silently become a
 * 28-INCH row (the footgun a defaulted unit invites). Positive only. */
function parseMeasureToTwips(raw: string): number | null {
	const match = raw.trim().match(/^(\d*\.?\d+)\s*(in|pt)$/i);
	if (!match) return null;
	const value = Number.parseFloat(match[1] as string);
	if (!Number.isFinite(value) || value <= 0) return null;
	const unit = (match[2] as string).toLowerCase();
	return Math.round(unit === "pt" ? value * 20 : value * 1440);
}

const VALIGNS = new Set(["top", "center", "bottom"]);
const HALIGNS = new Set(["left", "center", "right", "justify"]);
const TABLE_ALIGNS = new Set(["left", "center", "right"]);
const BORDER_STYLES = new Set(["single", "double", "none"]);
const HEIGHT_RULES = new Set(["atLeast", "exact", "auto"]);
const ALL_SIDES: CellBorderSide[] = [
	"top",
	"left",
	"bottom",
	"right",
	"insideH",
	"insideV",
];

const COLOR_NAMES: Record<string, string> = {
	white: "FFFFFF",
	black: "000000",
	gray: "808080",
	grey: "808080",
	lightgray: "D9D9D9",
	lightgrey: "D9D9D9",
	darkgray: "A6A6A6",
	darkgrey: "A6A6A6",
	red: "FF0000",
	green: "00B050",
	blue: "0070C0",
	yellow: "FFFF00",
};
