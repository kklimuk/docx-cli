import type { TableBorders, TableLayout } from "@core/table";
import { parseParagraphOptions, type RawValues } from "../insert/index";
import { parseTargetPlacement, placeSpec } from "../insert/place";
import {
	EXIT,
	fail,
	RENDER_VERIFY_EXAMPLE,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx tables create — insert a new (empty) table

Usage:
  docx tables create FILE --after LOCATOR --rows N --cols M [options]
  docx tables create FILE --before LOCATOR --rows N --cols M [options]
  docx tables create FILE (--at-start | --at-end) --rows N --cols M [options]

Examples:
  docx tables create doc.docx --after p3 --rows 3 --cols 2
  docx tables create doc.docx --after p3 --rows 2 --cols 3 --widths 1440,2880,4320
  docx tables create doc.docx --after p3 --rows 2 --cols 2 --table-width 50%
  docx tables create doc.docx --at-end --rows 4 --cols 3 --borders double
  # then fill cells from one read (batchable):
  docx edit doc.docx --at t0:r0c0 --text "Item"

Placement (exactly one required):
  --after LOCATOR   Insert after the block at LOCATOR (a pN / tN / cell paragraph)
  --before LOCATOR  Insert before the block at LOCATOR
  --at-start        Insert at the very top of the document (no locator needed)
  --at-end          Insert at the very end, before the trailing section properties

Grid (required):
  --rows N          Number of rows (>= 1)
  --cols N          Number of columns (>= 1)

Sizing (optional):
  --widths "A,B,C"  Column widths in twips, comma-separated; length must equal --cols
  --table-width V   Table total width, e.g. "100%" (default), "50%", or "4320" (twips)
  --borders S       single (default) | none | double
  --layout L        autofit (default; columns size to content) | fixed (honor
                    --widths exactly). Passing --widths implies fixed.

Paragraph options (apply to the table's anchor paragraph):
  --style NAME       Paragraph style      --alignment ALIGN  left|center|right|justify
  --space-before PT  --space-after PT     --line-spacing N
  --indent-left IN   --indent-right IN    --first-line IN   --hanging IN

Options:
  --track           Record the insert as a tracked change even when the
                    document's track-changes toggle is off. (Inserting a table
                    under tracking is rejected — Word has no honest construct.)
  --author NAME     Author for tracked changes (default: $DOCX_AUTHOR)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Report what would change without writing the file
  -v, --verbose     Print the success ack JSON (default: the minted locator)
  -h, --help        Show this help

The new table starts empty; fill each ordinary cell directly with \`docx edit
--at tN:rRcC --text "…"\` (batchable). Use \`:pK\` only for a complex cell or
exact paragraph targeting, then reshape with the other \`docx tables\` verbs
(insert-row, merge, set-widths, format, …).

AGENT VERIFICATION: \`docx read\` shows the grid but NOT how it lands on the
page. After adding a table, render and READ the images:
${RENDER_VERIFY_EXAMPLE}
Check the columns are sized sensibly (no content wrapping one char per line).

Output:
  Prints the new table's locator (tN). --verbose prints {ok:true, operation:
  "insert", …}. Errors print {code, error, hint?}.
`;

const OPTION_SPEC = {
	after: { type: "string" },
	before: { type: "string" },
	"at-start": { type: "boolean" },
	"at-end": { type: "boolean" },
	rows: { type: "string" },
	cols: { type: "string" },
	widths: { type: "string" },
	"table-width": { type: "string" },
	borders: { type: "string" },
	layout: { type: "string" },
	style: { type: "string" },
	alignment: { type: "string" },
	"space-before": { type: "string" },
	"space-after": { type: "string" },
	"line-spacing": { type: "string" },
	"indent-left": { type: "string" },
	"indent-right": { type: "string" },
	"first-line": { type: "string" },
	hanging: { type: "string" },
	author: { type: "string" },
	track: { type: "boolean" },
	...SAVE_FLAGS,
} as const;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(args, OPTION_SPEC, HELP);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const filePath = parsed.positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE argument", HELP);

	const placement = await parseTargetPlacement(parsed.values, HELP);
	if (typeof placement === "number") return placement;

	const tableFlags = await parseTableFlags(parsed.values);
	if (typeof tableFlags === "number") return tableFlags;

	const paragraphOptions = await parseParagraphOptions(parsed.values);
	if (typeof paragraphOptions === "number") return paragraphOptions;

	return placeSpec({
		filePath,
		placement,
		spec: { kind: "table", ...tableFlags },
		paragraphOptions,
		authorFlag: parsed.values.author as string | undefined,
		trackFlag: Boolean(parsed.values.track),
		outputPath: parsed.values.output as string | undefined,
		dryRun: Boolean(parsed.values["dry-run"]),
	});
}

/** Resolve `--rows`/`--cols` (required, >= 1) plus the optional sizing flags
 *  (`--widths`/`--table-width`/`--borders`/`--layout`) into the fields of a
 *  `table` spec. `--widths` are twips whose count must equal `--cols`;
 *  `--table-width` is a percentage ("50%", stored as OOXML fiftieths-of-a-pct)
 *  or twips; `--layout` defaults to `fixed` when `--widths` is set so the custom
 *  widths actually take effect. Moved here from `insert` when tables became a
 *  noun-verb — this is the sole owner of the table content spec. */
async function parseTableFlags(values: RawValues): Promise<
	| {
			rows: number;
			cols: number;
			widths?: number[];
			tableWidth?: { value: number; unit: "dxa" | "pct" };
			borders?: TableBorders;
			layout?: TableLayout;
	  }
	| number
> {
	const rowsRaw = values.rows as string | undefined;
	const colsRaw = values.cols as string | undefined;
	if (rowsRaw === undefined || colsRaw === undefined) {
		return fail("USAGE", "--rows and --cols are required", HELP);
	}
	const rows = Number.parseInt(rowsRaw, 10);
	const cols = Number.parseInt(colsRaw, 10);
	if (!Number.isFinite(rows) || rows < 1) {
		return fail("USAGE", `--rows must be a positive integer, got "${rowsRaw}"`);
	}
	if (!Number.isFinite(cols) || cols < 1) {
		return fail("USAGE", `--cols must be a positive integer, got "${colsRaw}"`);
	}

	const out: {
		rows: number;
		cols: number;
		widths?: number[];
		tableWidth?: { value: number; unit: "dxa" | "pct" };
		borders?: TableBorders;
		layout?: TableLayout;
	} = { rows, cols };

	const layoutRaw = values.layout as string | undefined;
	const widthsRaw = values.widths as string | undefined;
	if (widthsRaw !== undefined) {
		const widths = widthsRaw.split(",").map((part) => part.trim());
		const numeric: number[] = [];
		for (const part of widths) {
			const value = Number.parseInt(part, 10);
			if (!Number.isFinite(value) || value <= 0) {
				return fail(
					"USAGE",
					`--widths entries must be positive integers (twips), got "${part}"`,
				);
			}
			numeric.push(value);
		}
		if (numeric.length !== cols) {
			return fail(
				"USAGE",
				`--widths length (${numeric.length}) must equal --cols (${cols})`,
			);
		}
		out.widths = numeric;
	}

	const tableWidthRaw = values["table-width"] as string | undefined;
	if (tableWidthRaw !== undefined) {
		if (tableWidthRaw.endsWith("%")) {
			const pct = Number.parseFloat(tableWidthRaw.slice(0, -1));
			if (!Number.isFinite(pct) || pct <= 0 || pct > 100) {
				return fail(
					"USAGE",
					`--table-width percentage must be in (0, 100], got "${tableWidthRaw}"`,
				);
			}
			// OOXML pct units are fiftieths of a percent (5000 = 100%).
			out.tableWidth = { value: Math.round(pct * 50), unit: "pct" };
		} else {
			const twips = Number.parseInt(tableWidthRaw, 10);
			if (!Number.isFinite(twips) || twips <= 0) {
				return fail(
					"USAGE",
					`--table-width must be a positive integer (twips) or a percentage like "100%", got "${tableWidthRaw}"`,
				);
			}
			out.tableWidth = { value: twips, unit: "dxa" };
		}
	}

	const bordersRaw = values.borders as string | undefined;
	if (bordersRaw !== undefined) {
		if (
			bordersRaw !== "single" &&
			bordersRaw !== "double" &&
			bordersRaw !== "none"
		) {
			return fail(
				"USAGE",
				`--borders must be single, double, or none, got "${bordersRaw}"`,
			);
		}
		out.borders = bordersRaw === "single" ? "default" : { style: bordersRaw };
	}

	if (layoutRaw !== undefined) {
		if (layoutRaw !== "autofit" && layoutRaw !== "fixed") {
			return fail(
				"USAGE",
				`--layout must be autofit or fixed, got "${layoutRaw}"`,
			);
		}
		out.layout = layoutRaw;
	} else if (out.widths) {
		// Custom column widths are only honored under fixed layout — autofit
		// recomputes them from content. Default to fixed when --widths is given
		// so the widths actually take effect; an explicit --layout overrides.
		out.layout = "fixed";
	}

	return out;
}
