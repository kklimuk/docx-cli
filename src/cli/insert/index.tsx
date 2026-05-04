import {
	createRevisionAllocator,
	type DocView,
	Ins,
	isTrackChangesEnabled,
	markParagraphMarkAs,
	type Run,
	resolveAuthor,
	resolveDate,
	saveDocView,
	type TrackedMeta,
} from "@core";
import type { XmlNode } from "@core/parser";
import { parseArgs } from "util";
import {
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
	respond,
	writeStdout,
} from "../respond";
import { Paragraph, type ParagraphOptions } from "./emit";

const HELP = `docx insert — insert a paragraph at a locator

Usage:
  docx insert FILE [options]

Locator (one required):
  --after LOCATOR   Insert after the block at LOCATOR (e.g., p3)
  --before LOCATOR  Insert before the block at LOCATOR

Content (one required):
  --text TEXT       Insert a paragraph with this text
  --runs JSON       Insert a paragraph with custom runs (Run[] JSON)

Paragraph options:
  --style NAME       Apply paragraph style (e.g., Heading1)
  --alignment ALIGN  left | center | right | justify

Run options (only with --text):
  --color HEX       Run color, hex (e.g., 800080 for purple)
  --bold            Bold
  --italic          Italic

  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would be inserted; do not write the file
  -h, --help        Show this help

Examples:
  docx insert doc.docx --after p3 --text "Section header" --style Heading2
  docx insert doc.docx --before p0 --text "ALERT" --color CC0000 --bold
  docx insert doc.docx --after p2 --runs '[{"type":"text","text":"X","bold":true}]'
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				after: { type: "string" },
				before: { type: "string" },
				text: { type: "string" },
				runs: { type: "string" },
				style: { type: "string" },
				alignment: { type: "string" },
				color: { type: "string" },
				bold: { type: "boolean" },
				italic: { type: "boolean" },
				output: { type: "string", short: "o" },
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

	const after = parsed.values.after as string | undefined;
	const before = parsed.values.before as string | undefined;
	if (!after && !before) {
		return fail("USAGE", "Missing locator: pass --after or --before", HELP);
	}
	if (after && before) {
		return fail("USAGE", "Pass either --after or --before, not both", HELP);
	}

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

	const targetLocator = (after ?? before) as string;
	const blockRef = await resolveBlockOrFail(view, targetLocator);
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
	const insertIndex = after !== undefined ? targetIndex + 1 : targetIndex;

	if (isTrackChangesEnabled(view)) {
		applyTrackedInsertion(paragraphNode, view);
	}

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "insert",
			dryRun: true,
			path,
			locator: targetLocator,
			placement: after !== undefined ? "after" : "before",
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	blockRef.parent.splice(insertIndex, 0, paragraphNode);
	await saveDocView(view, outputPath);

	await respond({
		ok: true,
		operation: "insert",
		path: outputPath ?? path,
		locator: targetLocator,
		placement: after !== undefined ? "after" : "before",
	});
	return EXIT.OK;
}

function applyTrackedInsertion(paragraph: XmlNode, view: DocView): void {
	const allocator = createRevisionAllocator(view);
	const baseMeta = { author: resolveAuthor(), date: resolveDate() };
	const mintMeta = (): TrackedMeta => ({
		...baseMeta,
		revisionId: allocator.next(),
	});

	// Wrap each contiguous run of <w:r> children in a single <w:ins>, preserving
	// other children (e.g. <w:pPr>) at their existing positions.
	const newChildren: XmlNode[] = [];
	let runBuffer: XmlNode[] = [];
	const flush = (): void => {
		if (runBuffer.length === 0) return;
		newChildren.push(<Ins meta={mintMeta()}>{runBuffer}</Ins>);
		runBuffer = [];
	};
	for (const child of paragraph.children) {
		if (child.tag === "w:r") {
			runBuffer.push(child);
			continue;
		}
		flush();
		newChildren.push(child);
	}
	flush();
	paragraph.children = newChildren;

	// Mark the paragraph mark itself as inserted so accepting changes attributes
	// the new paragraph break.
	markParagraphMarkAs(paragraph, "ins", mintMeta());
}
