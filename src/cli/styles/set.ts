import { type Block, isBaselineStyle, iterateBlocks } from "@core";
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
import {
	hasStyleFormattingFlags,
	parseStyleFormatting,
	STYLE_FORMAT_FLAGS,
} from "./shared";

export const SET_HELP = `docx styles set — change an existing style's definition

Usage:
  docx styles set FILE --at STYLEID [formatting] [--name "Display Name"]
                                    [--based-on STYLEID] [--next STYLEID] [options]

Edits the style DEFINITION in word/styles.xml — every paragraph/run that uses
the style (and doesn't override the property locally) picks up the change. "Make
all Heading 1s green," "bump the body size to 12pt." For a one-off span, use
\`docx edit\`; to invent a new style, \`docx styles create\`.

The style edit is NOT tracked, even under track-changes — matching Word, which
applies style-definition edits to styles.xml directly (no redline). Paragraphs
with their OWN direct formatting keep it (the override wins); this never touches
the body, only the definition.

Run formatting (any style):
  --bold --italic --underline --strike --caps --smallcaps
  --superscript | --subscript
  --color HEX        Text color (e.g. 1F4E79 or #1F4E79)
  --font NAME        Font family (e.g. "Arial")
  --size PT          Font size in points (e.g. 16)
  --highlight NAME   Highlight color (yellow, green, cyan, …)
  --shade HEX        Background fill behind the text

Paragraph formatting (paragraph styles only):
  --alignment left|center|right|justify
  --space-before PT  --space-after PT   --line-spacing N|single|1.5|double
  --indent-left IN   --indent-right IN  --first-line IN | --hanging IN

Metadata:
  --name TEXT        Display name (the human-readable label)
  --based-on STYLEID Parent style to inherit from
  --next STYLEID     Style for the next paragraph

Options:
  -o, --output PATH  Write to PATH instead of overwriting FILE
  --dry-run          Print what would change; do not write the file
  -v, --verbose      Print the full success ack JSON
  -h, --help         Show this help

Examples:
  docx styles set report.docx --at Heading1 --color 1F4E79 --size 16 --bold
  docx styles set report.docx --at Normal --font "Times New Roman" --size 12
  docx styles set report.docx --at Quote --italic --indent-left 0.5 --space-after 6
`;

export async function runStylesSet(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			name: { type: "string" },
			"based-on": { type: "string" },
			next: { type: "string" },
			...STYLE_FORMAT_FLAGS,
			...SAVE_FLAGS,
		},
		SET_HELP,
	);
	if (typeof parsed === "number") return parsed;
	if (parsed.values.help) {
		await writeStdout(SET_HELP);
		return EXIT.OK;
	}
	setVerboseAck(Boolean(parsed.values.verbose));

	const { values } = parsed;
	const filePath = parsed.positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE argument", SET_HELP);
	const styleId = values.at as string | undefined;
	if (!styleId) {
		return fail(
			"USAGE",
			"Missing --at STYLEID (the style to edit)",
			"Run `docx styles FILE` to list style ids, or `docx styles create` to make a new one.",
		);
	}

	// An explicit empty metadata flag (`--name ""`) is treated as absent so we never
	// write a dangling `w:val=""`; `next` self-reference is valid (many styles follow
	// themselves) but a style based on ITSELF is an invalid inheritance cycle.
	const name = (values.name as string | undefined) || undefined;
	const basedOn = (values["based-on"] as string | undefined) || undefined;
	const next = (values.next as string | undefined) || undefined;
	const hasMeta =
		name !== undefined || basedOn !== undefined || next !== undefined;
	if (!hasStyleFormattingFlags(values) && !hasMeta) {
		return fail(
			"USAGE",
			"Nothing to set — pass a formatting flag (--bold/--color/--size/…) or metadata (--name/--based-on/--next)",
			SET_HELP,
		);
	}
	if (basedOn === styleId) {
		return fail(
			"USAGE",
			`A style can't be based on itself ("${styleId}")`,
			"Point --based-on at a different (parent) style, or drop it.",
		);
	}

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	const styles = document.ensureStyles();
	// Auto-provision an un-materialized baseline (e.g. `set --at Heading1` on a doc
	// that never used Heading1) so the agent's intent — "make Heading1 look like X"
	// — works regardless of whether the def already exists. A non-baseline unknown
	// id is a real error (no definition to edit and no template to provision).
	if (!styles.hasStyle(styleId)) {
		if (isBaselineStyle(styleId)) {
			styles.ensureStyle(styleId);
		} else {
			return fail(
				"BLOCK_NOT_FOUND",
				`No style with id "${styleId}"`,
				"Run `docx styles FILE` to list style ids, or `docx styles create` to make a new one.",
			);
		}
	}

	// The real `w:type` (paragraph/character/table/numbering) — reported honestly in
	// the ack and used to gate paragraph-only flags. Default to paragraph only when
	// the attribute is genuinely absent.
	const styleType =
		styles.getStyle(styleId)?.getAttribute("w:type") || "paragraph";
	const formatting = parseStyleFormatting(values, styleType);
	if ("error" in formatting) {
		return fail("USAGE", formatting.error, formatting.hint);
	}

	styles.setStyleFormatting(styleId, {
		runFormat: formatting.runFormat,
		paragraphOptions: formatting.paragraphOptions,
		name,
		basedOn,
		next,
	});

	const usage = countStyleUsage(document, styleId);
	const dryRun = Boolean(parsed.values["dry-run"]);
	const outputPath = values.output as string | undefined;

	if (dryRun) {
		await respond({
			operation: "styles.set",
			dryRun: true,
			id: styleId,
			type: styleType,
			usedBy: usage,
		});
		return EXIT.OK;
	}

	await document.save(outputPath);

	// Surface the blast radius: how many places use the style, and the reminder
	// that a paragraph's own direct formatting still wins (we don't strip it).
	const note =
		usage > 0
			? `${usage === 1 ? "1 place uses" : `${usage} places use`} this style. Text with its own direct formatting keeps it.`
			: "No body content uses this style yet.";

	await respondAck(
		{
			ok: true,
			operation: "styles.set",
			path: outputPath ?? filePath,
			id: styleId,
			type: styleType,
			usedBy: usage,
		},
		note,
	);
	return EXIT.OK;
}

/** Count the body references to `styleId` — paragraph `pStyle`, run `rStyle`, and
 *  table `tblStyle` (descending into cells via `iterateBlocks`) — for the ack's
 *  blast-radius note. */
function countStyleUsage(
	document: { body: { blocks: Block[] } },
	styleId: string,
): number {
	let count = 0;
	for (const block of iterateBlocks(document.body.blocks)) {
		if (block.type === "table") {
			if (block.style === styleId) count++;
			continue;
		}
		if (block.type !== "paragraph") continue;
		if (block.style === styleId) count++;
		for (const run of block.runs) {
			if (run.type === "text" && run.runStyle === styleId) count++;
		}
	}
	return count;
}
