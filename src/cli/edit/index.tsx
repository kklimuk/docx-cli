import { type Run, saveDocView } from "@core";
import type { XmlNode } from "@core/parser";
import { parseArgs } from "util";
import { Paragraph, type ParagraphOptions } from "../insert/emit";
import {
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
	respond,
	writeStdout,
} from "../respond";

const HELP = `docx edit — replace a paragraph at a locator

Usage:
  docx edit FILE [options]

Locator (required):
  --at LOCATOR      Block to replace (e.g., p3)

Content (one required):
  --text TEXT       Replace with a single-run paragraph
  --runs JSON       Replace with custom runs (Run[] JSON)

Paragraph options:
  --style NAME       Paragraph style (e.g., Heading1)
  --alignment ALIGN  left | center | right | justify

Run options (only with --text):
  --color HEX       Run color, hex (e.g., 800080 for purple)
  --bold            Bold
  --italic          Italic

  --dry-run         Print what would change; do not write the file
  -h, --help        Show this help

Examples:
  docx edit doc.docx --at p3 --text "Replaced." --style Heading2
  docx edit doc.docx --at p0 --runs '[{"type":"text","text":"X","bold":true}]'
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				at: { type: "string" },
				text: { type: "string" },
				runs: { type: "string" },
				style: { type: "string" },
				alignment: { type: "string" },
				color: { type: "string" },
				bold: { type: "boolean" },
				italic: { type: "boolean" },
				"dry-run": { type: "boolean" },
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

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const at = parsed.values.at as string | undefined;
	if (!at) return fail("USAGE", "Missing --at LOCATOR", HELP);

	const text = parsed.values.text as string | undefined;
	const runsJson = parsed.values.runs as string | undefined;
	if (!text && !runsJson) {
		return fail("USAGE", "Missing content: pass --text or --runs", HELP);
	}
	if (text && runsJson) {
		return fail("USAGE", "Pass either --text or --runs, not both", HELP);
	}

	const paragraphOptions: ParagraphOptions = {};
	const styleValue = parsed.values.style as string | undefined;
	if (styleValue) paragraphOptions.style = styleValue;
	const alignmentValue = parsed.values.alignment as string | undefined;
	if (alignmentValue) {
		if (
			alignmentValue !== "left" &&
			alignmentValue !== "center" &&
			alignmentValue !== "right" &&
			alignmentValue !== "justify"
		) {
			return fail(
				"USAGE",
				`Invalid --alignment: ${alignmentValue}`,
				"Valid values: left, center, right, justify",
			);
		}
		paragraphOptions.alignment = alignmentValue;
	}

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const blockRef = await resolveBlockOrFail(view, at);
	if (typeof blockRef === "number") return blockRef;

	let paragraphNode: XmlNode;
	if (text !== undefined) {
		const color = parsed.values.color as string | undefined;
		paragraphNode = (
			<Paragraph
				text={text}
				{...paragraphOptions}
				{...(color ? { color } : {})}
				{...(parsed.values.bold ? { bold: true as const } : {})}
				{...(parsed.values.italic ? { italic: true as const } : {})}
			/>
		);
	} else {
		let runsValue: Run[];
		try {
			runsValue = JSON.parse(runsJson as string) as Run[];
		} catch (jsonError) {
			const message =
				jsonError instanceof Error ? jsonError.message : String(jsonError);
			return fail("USAGE", `Invalid --runs JSON: ${message}`);
		}
		if (!Array.isArray(runsValue)) {
			return fail("USAGE", "--runs must be a JSON array of Run objects");
		}
		paragraphNode = <Paragraph runs={runsValue} {...paragraphOptions} />;
	}

	const targetIndex = blockRef.parent.indexOf(blockRef.node);
	if (targetIndex === -1) {
		return fail(
			"BLOCK_NOT_FOUND",
			"Block reference is stale (parent does not contain it)",
		);
	}

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "edit",
			dryRun: true,
			path,
			locator: at,
		});
		return EXIT.OK;
	}

	blockRef.parent.splice(targetIndex, 1, paragraphNode);
	await saveDocView(view);

	await respond({ ok: true, operation: "edit", path, locator: at });
	return EXIT.OK;
}
