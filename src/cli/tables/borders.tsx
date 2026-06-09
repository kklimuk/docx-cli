import { describeForms } from "@core";
import { w } from "@core/jsx";
import { parseTableAt } from "@core/locators";
import type { XmlNode } from "@core/parser";
import {
	buildGrid,
	resolveTableNode,
	setTablePropertiesChild,
} from "@core/table";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";
import { noteStructuralChange } from "./common";

const STYLES = new Set(["single", "double", "none"]);

const AT_FORMS = describeForms(["table"], "                     ");

const HELP = `docx tables borders — set table borders

Usage:
  docx tables borders FILE --at tN [options]

Required:
  --at LOCATOR       Target table. Supports:
${AT_FORMS}
                     See \`docx info locators\`.

Optional:
  --style STYLE      single | double | none (default: single)
  --size EIGHTHS     Border thickness in eighths of a point (default: 4 = 0.5pt)
  --color HEX        Hex color without '#', or "auto" (default: auto)
  --author NAME      Author for the audit comment when track-changes is on
                     (default: $DOCX_AUTHOR)
  -o, --output PATH  Write to PATH instead of overwriting FILE
  --dry-run          Print what would change; do not write
  -v, --verbose      Print the success ack JSON
  -h, --help         Show this help

Applies to all six table border edges (<w:tblBorders>). OOXML has no
tracked-change construct Word will round-trip for a hand-authored border change
(Word does not revert a tblPrChange we author on reject), so under track-changes
the change is applied in place with a [docx-cli] audit comment instead.

Output:
  Silent on success (exit 0). --verbose prints {ok:true, operation, path, table,
  style, ...}. --dry-run prints the preview object (no ok field). Errors print
  {code, error, hint?} with a nonzero exit.

Examples:
  docx tables borders doc.docx --at t0 --style double --size 8 --color 444444
  docx tables borders doc.docx --at t0 --style none
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			style: { type: "string" },
			size: { type: "string" },
			color: { type: "string" },
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

	const style = (parsed.values.style as string | undefined) ?? "single";
	if (!STYLES.has(style)) {
		return fail("USAGE", "--style must be single, double, or none");
	}
	const size = parsed.values.size as string | undefined;
	if (
		size !== undefined &&
		(!Number.isInteger(Number(size)) || Number(size) <= 0)
	) {
		return fail(
			"USAGE",
			"--size must be a positive integer (eighths of a point)",
		);
	}
	const sizeEighths = size ? Number(size) : 4;
	const color = (parsed.values.color as string | undefined) ?? "auto";

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const tableNode = resolveTableNode(document, tableId);
	if (!tableNode) return fail("BLOCK_NOT_FOUND", `Table not found: ${tableId}`);

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			operation: "tables.borders",
			dryRun: true,
			path,
			table: tableId,
			style,
			...(style !== "none" ? { size: sizeEighths, color } : {}),
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	setTablePropertiesChild(
		tableNode,
		"w:tblBorders",
		buildBorders(style, sizeEighths, color),
	);

	// Word does not record a table border change as a tracked revision that it
	// will revert on reject (verified: a tblPrChange we author isn't reverted by
	// Word's reject), so we apply it in place and note it with an audit comment.
	noteStructuralChange(
		document,
		buildGrid(tableNode).rows[0]?.cells[0]?.node,
		`table borders set (${style})`,
		parsed.values.author as string | undefined,
		parsed.values.track as boolean | undefined,
	);

	await document.save(outputPath);

	await respondAck({
		ok: true,
		operation: "tables.borders",
		path: outputPath ?? path,
		table: tableId,
		style,
		...(style !== "none" ? { size: sizeEighths, color } : {}),
	});
	return EXIT.OK;
}

function buildBorders(
	style: string,
	sizeEighths: number,
	color: string,
): XmlNode {
	const attrs =
		style === "none"
			? { "w-val": "none" }
			: {
					"w-val": style,
					"w-sz": String(sizeEighths),
					"w-space": "0",
					"w-color": color,
				};
	return (
		<w.tblBorders>
			<w.top {...attrs} />
			<w.left {...attrs} />
			<w.bottom {...attrs} />
			<w.right {...attrs} />
			<w.insideH {...attrs} />
			<w.insideV {...attrs} />
		</w.tblBorders>
	);
}
