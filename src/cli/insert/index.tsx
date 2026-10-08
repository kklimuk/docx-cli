import { describeForms, type InsertSpec } from "@core";
import type { ParagraphOptions } from "@core/blocks";
import type { parseArgs } from "util";
import {
	adoptPositionalMarkdown,
	batchExampleIntro,
	decodeInlineEscapes,
	parseRunsArg,
	parseSpacingIndentFlags,
	pickContextualHelp,
	rejectShellMangledValue,
} from "../parse-helpers";
import {
	EXIT,
	fail,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";
import { runInsertBatch } from "./batch";
import { parseTargetPlacement, placeSpec, type TargetPlacement } from "./place";

const ANCHOR_FORMS = describeForms(
	[
		"paragraph",
		"table",
		"section",
		"cell",
		"cellParagraph",
		"textBoxParagraph",
	],
	"                      ",
);

const INSERT_HELP = `docx insert — insert content at a locator

The read view is Markdown; add new content the same way — just pass it after the
locator. It's parsed (headings, lists, tables, links, bold, inline <span>/<mark>)
in the same dialect \`read\` prints, and a multi-block source inserts several
blocks at once. --text is the LITERAL escape hatch: every character lands
verbatim (--text "**bold**" writes the asterisks). New content blends in: a
paragraph takes its neighbor's font/size, tab stops, indent, alignment and
spacing; a "- item" next to a list joins THAT list with its look; anything else
falls back to the document's font/size (the docx:base \`read\` shows).

Usage:
  docx insert FILE (--after | --before | --at) LOCATOR "MARKDOWN" [options]
  docx insert FILE (--at-start | --at-end) "MARKDOWN" [options]
  docx insert FILE --batch FILE.jsonl [options]   # many inserts, one read (- = stdin)

Examples:
${batchExampleIntro("Insert several blocks")}
  #   adds.jsonl:
  #     {"after":"p3","markdown":"New clause."}
  #     {"at":"t0:r2c1","markdown":"Charlie Darwin"}            # fill a blank cell
  #     {"after":"p5","markdown":"## Summary\\n\\nFirst point."}  # several blocks
  #     {"before":"p0","markdown":"**ALERT**","alignment":"center"}
  docx insert doc.docx --batch adds.jsonl
  # …or one at a time:
  docx insert doc.docx --after p3 "## New section"
  docx insert doc.docx --after p3 "- first\\n- second"   # a bullet list
  docx insert doc.docx --after p7 "- another bullet"      # p7 is a bullet: joins its list
  docx insert doc.docx --after p3 "See [the site](https://example.com)."
  docx insert doc.docx --after p9 "Harvard SEAS\\tCambridge, MA"  # inherits p9's tabs
  docx insert doc.docx --at-start "# Title"
  docx insert doc.docx --after p3 --page-break
  docx insert doc.docx --after p3 --text-file notes.txt   # literal prose, one paragraph per line

Ordering: batch entries apply in file order; several anchored after the SAME
block stack in that order (three "after":"p0" land as p1, p2, p3).

Placement (exactly one):
  --at LOCATOR      After an ordinary block; INTO a bare table cell (fills a blank
                    cell, otherwise appends).
  --after LOCATOR   After a block, or at the END of a bare cell
  --before LOCATOR  Before a block, or at the START of a bare cell
                    LOCATOR is one of:
${ANCHOR_FORMS}
  --at-start / --at-end   Very top / very end of the document (not in --batch).

Content (one):
  "MARKDOWN"        Positional, after the flags: parsed GFM, the dialect \`read\`
                    prints. Quote it as ONE argument ('$' amounts: single quotes).
  --markdown TEXT   The same, as a flag. --markdown-file PATH reads it from a
                    file ("-" = stdin).
  --text TEXT       LITERAL single paragraph, no parsing. Format it with
                    --bold/--italic/--color/--url. \`docx insert --text --help\`.
  --text-file PATH  LITERAL multi-paragraph text ("-" = stdin): every character
                    verbatim, each newline a new paragraph. For prose Markdown
                    would mangle ("3. note", bare URLs, *x*, {++x++}).
  --runs JSON       Byte-precise runs (Run[] JSON). \`docx insert --runs --help\`.
  --page-break / --column-break   An empty paragraph holding that break.

Formatting (applies to every paragraph the content produces):
  --style NAME       Paragraph style — refused next to --markdown (a \`# heading\`
                     or list item sets its own)
  --alignment ALIGN  left | center | right | justify
  --space-before N / --space-after N   Points, or a unit suffix (6pt, 0.1in, 120tw)
  --line-spacing N   1, 1.5, 2, single, double, or 15pt
  --indent-left N / --indent-right N   Inches, or a unit suffix (0.5in, 1.27cm,
                     720tw); a bare number above 9 is read as twips
  --first-line N / --hanging N         Same units
  --list KIND / --list-level N   Make a --text/--runs paragraph a bullet|ordered
                     item at nesting level 0-8 (in markdown, write "- item").

Batch (--batch PATH | -): one JSON object per line; keys mirror the flags
  ({"at"|"after"|"before":…} + one content field + any formatting). Locators
  address the document AS READ. --at-start/--at-end are not batchable.

Options:
  --track           Record as a tracked change even when the doc toggle is off
  --author NAME     Tracked-change author (default: $DOCX_AUTHOR)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Preview; write nothing
  -v, --verbose     Full JSON ack
  -h, --help        Show this help

Output: the new block's locator(s), one per line (exit 0). Ids shift after an
insert — re-read before further edits, or do everything from one read with
--batch. Errors print {code, error, hint?} with a nonzero exit.`;

const INSERT_TEXT_HELP = `docx insert --text — insert new text content and format it

Usage:
  docx insert FILE --at LOCATOR --text "New paragraph" [options]
  docx insert FILE (--after | --before) LOCATOR --text "New paragraph" [options]
  docx insert FILE --at-start --text "Title" --style Title

Examples:
  docx insert doc.docx --at t0:r2c1 --text "Charlie Darwin"
  docx insert doc.docx --at p3 --text "Paragraph after p3" --style Heading2
  docx insert doc.docx --after p3 --text "Section header" --style Heading2
  docx insert doc.docx --before p0 --text "ALERT" --color CC0000 --bold
  docx insert doc.docx --after p3 --text "click here" --url https://example.com
  docx insert doc.docx --after p3 --markdown "A **bold** intro line."

--text builds a NEW paragraph from LITERAL characters. A markdown-looking
value (e.g. **bold**) will insert the literal **bold** characters. To get formatting:

  Ride-along run formatting (formats the whole new run):
    --bold            Bold
    --italic          Italic
    --color HEX       Run color (e.g. CC0000 — no '#')
    --url URL         Wrap the inserted text in a hyperlink to URL
  e.g. \`--after pN --text "click here" --url https://example.com --bold\`.

  For richer / MIXED formatting (**bold**, \`code\`, [links](url), lists, headings),
  use --markdown instead of --text:
    docx insert FILE --after pN --markdown "A **bold** word and a [link](url)."

  For exact per-run control, use --runs (Run[] JSON) — see \`docx insert --runs --help\`.

Paragraph options ride along too: --style, --alignment, --list bullet|ordered,
--space-*/--line-spacing/--indent-*. (These do NOT combine with --markdown.)

Literal bulk prose — use --text-file PATH ("-" = stdin): every character verbatim,
each newline a new paragraph, no GFM parsing. The safe channel for prose with
"3." lists, bare URLs, *x*, {++x++} that GFM would otherwise corrupt.
`;

const INSERT_RUNS_HELP = `docx insert --runs — insert a paragraph from explicit runs (Run[] JSON)

Examples:
  docx insert doc.docx --after p2 --runs '[{"type":"text","text":"X","bold":true}]'
  docx insert doc.docx --after p2 --runs '[{"type":"text","text":"H","size":12},{"type":"text","text":"2","vertAlign":"subscript"},{"type":"text","text":"O"}]'

--runs JSON builds a NEW paragraph from an array of runs. Each run object may carry:
  { "type": "text", "text": "…",
    "bold": true, "italic": true, "underline": true, "strike": true,
    "color": "C00000",         // hex, no '#'
    "highlight": "yellow",     // named highlighter
    "shade": "EEEEEE",         // background fill, hex
    "font": "Times New Roman", "size": 12,
    "caps": true, "smallcaps": true,
    "vertAlign": "superscript" | "subscript" }
  e.g. --runs '[{"type":"text","text":"Note: ","bold":true},{"type":"text","text":"see clause 4."}]'

Prefer --text (with --bold/--italic/--color/--url) or --markdown unless you need
exact per-run control — --runs is the escape hatch when one line mixes fonts,
sizes, super/subscript, or highlight/shade the simpler flags can't express.

Paragraph options (--style/--alignment/--list/--space-*/…) ride along with --runs
just like --text. To FORMAT text that already EXISTS (not insert new), use \`docx
edit\` — see \`docx edit --runs --help\`.
`;

export async function run(args: string[]): Promise<number> {
	const help = pickContextualHelp(args, {
		default: INSERT_HELP,
		text: INSERT_TEXT_HELP,
		runs: INSERT_RUNS_HELP,
	});
	const parsed = await tryParseArgs(args, OPTION_SPEC, help);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(help);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const filePath = parsed.positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE argument", INSERT_HELP);
	const positionalError = await adoptPositionalMarkdown(parsed, INSERT_HELP);
	if (positionalError !== undefined) return positionalError;

	const batchInput = parsed.values.batch as string | undefined;
	if (batchInput !== undefined) {
		return runInsertBatch(filePath, batchInput, parsed.values);
	}

	const opts = await buildSingleShotOptions(filePath, parsed.values);
	if (typeof opts === "number") return opts;

	return placeSpec(opts);
}

async function buildSingleShotOptions(
	filePath: string,
	values: RawValues,
): Promise<ValidatedOptions | number> {
	const placement = await parseTargetPlacement(values, INSERT_HELP, {
		allowAt: true,
	});
	if (typeof placement === "number") return placement;

	const spec = await chooseContentSpec(values);
	if (typeof spec === "number") return spec;

	// `--text` writes literal characters — a markdown-looking value (e.g. **bold**)
	// lands verbatim, by design (use --markdown to parse it). We still refuse a
	// shell-gutted currency value ("$300" → ".00"), which is never intentional.
	if (spec.kind === "text") {
		const mangled = await rejectShellMangledValue(spec.text, "--text");
		if (typeof mangled === "number") return mangled;
	}

	// A markdown source owns its paragraph style and list membership (heading
	// levels, list numbering), so `--style`/`--list` conflict — reject them up
	// front. Layout flags ride along (see MARKDOWN_INCOMPATIBLE_FLAGS).
	if (spec.kind === "markdown") {
		const conflict = MARKDOWN_INCOMPATIBLE_FLAGS.find(
			(flag) => values[flag] !== undefined,
		);
		if (conflict) {
			return fail(
				"USAGE",
				`--${conflict} can't be combined with markdown content (positional, --markdown, or --markdown-file) — the markdown source controls block-level styling`,
				INSERT_HELP,
			);
		}
	}

	const paragraphOptions = await parseParagraphOptions(values);
	if (typeof paragraphOptions === "number") return paragraphOptions;

	return {
		filePath,
		placement,
		spec,
		paragraphOptions,
		authorFlag: values.author as string | undefined,
		trackFlag: Boolean(values.track),
		outputPath: values.output as string | undefined,
		dryRun: Boolean(values["dry-run"]),
		allowCellTarget: true,
	};
}

const OPTION_SPEC = {
	at: { type: "string" },
	after: { type: "string" },
	before: { type: "string" },
	"at-start": { type: "boolean" },
	"at-end": { type: "boolean" },
	batch: { type: "string" },
	text: { type: "string" },
	"text-file": { type: "string" },
	runs: { type: "string" },
	"page-break": { type: "boolean" },
	"column-break": { type: "boolean" },
	section: { type: "boolean" },
	columns: { type: "string" },
	type: { type: "string" },
	table: { type: "boolean" },
	rows: { type: "string" },
	cols: { type: "string" },
	widths: { type: "string" },
	"table-width": { type: "string" },
	borders: { type: "string" },
	layout: { type: "string" },
	image: { type: "string" },
	alt: { type: "string" },
	width: { type: "string" },
	height: { type: "string" },
	caption: { type: "string" },
	code: { type: "string" },
	"code-file": { type: "string" },
	language: { type: "string" },
	task: { type: "string" },
	list: { type: "string" },
	"list-level": { type: "string" },
	equation: { type: "string" },
	display: { type: "boolean" },
	markdown: { type: "string" },
	"markdown-file": { type: "string" },
	style: { type: "string" },
	alignment: { type: "string" },
	"space-before": { type: "string" },
	"space-after": { type: "string" },
	"line-spacing": { type: "string" },
	"indent-left": { type: "string" },
	"indent-right": { type: "string" },
	"first-line": { type: "string" },
	hanging: { type: "string" },
	color: { type: "string" },
	bold: { type: "boolean" },
	italic: { type: "boolean" },
	url: { type: "string" },
	author: { type: "string" },
	track: { type: "boolean" },
	...SAVE_FLAGS,
} as const;

type ValidatedOptions = {
	filePath: string;
	placement: TargetPlacement;
	spec: InsertSpec;
	paragraphOptions: ParagraphOptions;
	authorFlag?: string;
	trackFlag: boolean;
	outputPath?: string;
	dryRun: boolean;
	allowCellTarget: true;
};

export type RawValues = ReturnType<typeof parseArgs>["values"];

/** The only flags that CONFLICT with `--markdown`: a `# heading` or list item in
 *  the source owns its paragraph style and list membership. Layout flags ride
 *  along (`applyParagraphOptionsToBlocks`). */
export const MARKDOWN_INCOMPATIBLE_FLAGS = [
	"style",
	"list",
	"list-level",
] as const;

/** The mutually-exclusive content flags, each with the sub-flags that only
 * make sense alongside it. Drives both the "exactly one content flag" check
 * and the "this sub-flag requires its content flag" check, so those rules
 * live in one place instead of scattered guards. */
const CONTENT_KINDS = [
	{ flag: "text", subFlags: ["color", "bold", "italic", "url"] },
	{ flag: "text-file", subFlags: [] },
	{ flag: "runs", subFlags: [] },
	{ flag: "page-break", subFlags: [] },
	{ flag: "column-break", subFlags: [] },
	{ flag: "markdown", subFlags: [] },
	{ flag: "markdown-file", subFlags: [] },
] as const;

const CONTENT_FLAG_LIST = CONTENT_KINDS.map((kind) => `--${kind.flag}`).join(
	", ",
);

export async function chooseContentSpec(
	values: RawValues,
): Promise<InsertSpec | number> {
	// `insert` no longer creates sections/columns. A raw section break formats the
	// content ABOVE it (the off-by-one that traps weak agents); `docx sections`
	// takes a range and inserts the bounding breaks so the columns land exactly
	// where you name them. Redirect rather than silently ignore the flags.
	if (
		values.section !== undefined ||
		values.columns !== undefined ||
		values.type !== undefined
	) {
		return fail(
			"USAGE",
			"insert no longer creates section/column layout — use `docx sections`",
			"To put paragraphs pN…pM in N columns: `docx sections --at pN-pM --columns N`. To recount an existing section: `docx sections --at sN --columns N`.",
		);
	}
	// Code blocks moved to their own noun-verb command so insert stays lean.
	if (
		values.code !== undefined ||
		values["code-file"] !== undefined ||
		values.language !== undefined
	) {
		return fail(
			"USAGE",
			"insert no longer builds code blocks — use `docx code add`",
			"e.g. `docx code add FILE --after pN --code-file snippet.py --language python`. See `docx code add --help`.",
		);
	}
	// Equations moved to their own noun-verb command too.
	if (values.equation !== undefined || values.display !== undefined) {
		return fail(
			"USAGE",
			"insert no longer builds equations — use `docx equations add`",
			'e.g. `docx equations add FILE --after pN --equation "x^2 + y^2" --display`. See `docx equations add --help`.',
		);
	}
	// Task-list checkboxes moved to their own noun-verb command too. (`--list`
	// bullet/ordered still lives here — only the checkbox variant moved.)
	if (values.task !== undefined) {
		return fail(
			"USAGE",
			"insert no longer builds task-list items — use `docx tasks add`",
			'e.g. `docx tasks add FILE --after pN --text "buy groceries" --checked` (or --unchecked). See `docx tasks add --help`.',
		);
	}
	// Images moved to their own noun-verb command too. Redirect on --image or any
	// of its sub-flags (none of which is shared by another content kind).
	if (
		values.image !== undefined ||
		values.alt !== undefined ||
		values.width !== undefined ||
		values.height !== undefined ||
		values.caption !== undefined
	) {
		return fail(
			"USAGE",
			"insert no longer builds images — use `docx images add`",
			'e.g. `docx images add FILE --after pN --image chart.png --alt "Figure 1"`. See `docx images add --help`.',
		);
	}
	// Tables moved to their own noun-verb command too. Redirect on --table or any
	// of its sub-flags (all table-exclusive — images use --width/--height, not
	// --widths). Keep the flags in OPTION_SPEC so this fires instead of erroring
	// on an unknown flag.
	if (
		values.table !== undefined ||
		values.rows !== undefined ||
		values.cols !== undefined ||
		values.widths !== undefined ||
		values["table-width"] !== undefined ||
		values.borders !== undefined ||
		values.layout !== undefined
	) {
		return fail(
			"USAGE",
			"insert no longer builds tables — use `docx tables create`",
			"e.g. `docx tables create FILE --after pN --rows 3 --cols 2`. See `docx tables create --help`.",
		);
	}
	const present = CONTENT_KINDS.filter(
		(kind) => values[kind.flag] !== undefined,
	);
	if (present.length > 1) {
		return fail("USAGE", `Pass only one of ${CONTENT_FLAG_LIST}`, INSERT_HELP);
	}
	const chosen = present[0];
	if (!chosen) {
		return fail(
			"USAGE",
			`Missing content: pass ${CONTENT_FLAG_LIST}`,
			INSERT_HELP,
		);
	}

	// Reject sub-flags belonging to a content kind other than the chosen one,
	// so e.g. `--columns` without `--section` is an error rather than ignored.
	// A subFlag listed under MULTIPLE kinds (e.g. `--language` shared by both
	// `--code` and `--code-file`) is permitted if the chosen kind is one of
	// them — only orphans wholly unrelated to the chosen kind error.
	const chosenSubFlags = new Set<string>(chosen.subFlags);
	for (const kind of CONTENT_KINDS) {
		if (kind.flag === chosen.flag) continue;
		const orphan = kind.subFlags.find(
			(flag) => values[flag] !== undefined && !chosenSubFlags.has(flag),
		);
		if (orphan) {
			return fail("USAGE", `--${orphan} requires --${kind.flag}`, INSERT_HELP);
		}
	}

	switch (chosen.flag) {
		case "text":
			return buildTextSpec(values);
		case "text-file":
			return resolveLiteralSpec(values);
		case "runs": {
			const runs = await parseRunsArg(values.runs as string);
			return typeof runs === "number" ? runs : { kind: "runs", runs };
		}
		case "page-break":
			return { kind: "break", breakKind: "page" };
		case "column-break":
			return { kind: "break", breakKind: "column" };
		case "markdown":
		case "markdown-file":
			return resolveMarkdownSpec(values, chosen.flag);
	}
}

/** Resolve `--markdown TEXT` (inline) or `--markdown-file PATH` (file / stdin)
 *  into a uniform `markdown` spec. Stdin path mirrors `--code-file -`. */
async function resolveMarkdownSpec(
	values: RawValues,
	flag: "markdown" | "markdown-file",
): Promise<Extract<InsertSpec, { kind: "markdown" }> | number> {
	if (flag === "markdown") {
		const source = decodeInlineEscapes(values.markdown as string);
		return { kind: "markdown", source };
	}
	const path = values["markdown-file"] as string;
	try {
		const source =
			path === "-"
				? await new Response(Bun.stdin.stream()).text()
				: await Bun.file(path).text();
		return { kind: "markdown", source };
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail(
			"FILE_NOT_FOUND",
			`Failed to read --markdown-file ${path}: ${message}`,
		);
	}
}

/** Resolve `--text-file PATH` (file / stdin) into a `literal` spec — the
 *  parser-free channel. Every newline in the file starts a new paragraph and
 *  every other character lands verbatim (no GFM parsing), so reviewer prose
 *  with `3. …`, `*x*`, `[t](u)`, bare URLs, `{++x++}` survives untouched. */
async function resolveLiteralSpec(
	values: RawValues,
): Promise<Extract<InsertSpec, { kind: "literal" }> | number> {
	const path = values["text-file"] as string;
	try {
		const text =
			path === "-"
				? await new Response(Bun.stdin.stream()).text()
				: await Bun.file(path).text();
		return { kind: "literal", text };
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail(
			"FILE_NOT_FOUND",
			`Failed to read --text-file ${path}: ${message}`,
		);
	}
}

function buildTextSpec(
	values: RawValues,
): Extract<InsertSpec, { kind: "text" }> {
	const url = values.url as string | undefined;
	return {
		kind: "text",
		text: decodeInlineEscapes(values.text as string),
		format: {
			color: values.color as string | undefined,
			bold: values.bold as boolean | undefined,
			italic: values.italic as boolean | undefined,
		},
		...(url ? { hyperlinkUrl: url } : {}),
	};
}

export async function parseParagraphOptions(
	values: RawValues,
): Promise<ParagraphOptions | number> {
	const out: ParagraphOptions = {};

	const styleValue = values.style as string | undefined;
	if (styleValue) out.style = styleValue;

	const alignmentValue = values.alignment as string | undefined;
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
		out.alignment = alignmentValue;
	}

	const listValue = values.list as string | undefined;
	const listLevelValue = values["list-level"] as string | undefined;

	if (listValue !== undefined) {
		if (listValue !== "bullet" && listValue !== "ordered") {
			return fail(
				"USAGE",
				`--list must be "bullet" or "ordered", got "${listValue}"`,
				INSERT_HELP,
			);
		}
		// Mark the intent to allocate a list; the numId is resolved later in
		// `resolveListContext` (post-document-open) using the same anchor-inherit
		// logic a task item uses (`tasks add` sets `taskState` and reuses this
		// resolver). We stash the kind on a side channel so the resolver knows
		// which abstractNum to use.
		out.list = { level: 0, numId: -1 };
		(out as ParagraphOptions & { listKind?: "bullet" | "ordered" }).listKind =
			listValue;
	}

	if (listLevelValue !== undefined) {
		const level = Number(listLevelValue);
		if (!Number.isInteger(level) || level < 0 || level > 8) {
			return fail(
				"USAGE",
				`--list-level must be an integer 0-8, got "${listLevelValue}"`,
				INSERT_HELP,
			);
		}
		if (out.list) out.list.level = level;
		// If --list isn't set (e.g. a `tasks add` item, or inheritance), we still
		// record the level — it applies once the resolver attaches a list.
		(out as ParagraphOptions & { explicitLevel?: number }).explicitLevel =
			level;
	}

	const spacingIndent = parseSpacingIndentFlags(values);
	if ("error" in spacingIndent) {
		return fail("USAGE", spacingIndent.error, spacingIndent.hint);
	}
	if (spacingIndent.spacing) out.spacing = spacingIndent.spacing;
	if (spacingIndent.indent) out.indent = spacingIndent.indent;

	return out;
}
