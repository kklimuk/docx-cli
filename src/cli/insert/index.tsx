import { describeForms, type InsertSpec } from "@core";
import type { ParagraphOptions } from "@core/blocks";
import type { parseArgs } from "util";
import {
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
	["paragraph", "table", "section", "cellParagraph"],
	"                      ",
);

const INSERT_HELP = `docx insert — insert content at a locator

Usage:
  docx insert FILE (--after | --before) LOCATOR <content> [options]
  docx insert FILE (--at-start | --at-end) <content> [options]
  docx insert FILE --batch FILE.jsonl [options]   # many inserts, one read
  docx insert FILE --batch -          [options]   # read JSONL from stdin

Examples:
${batchExampleIntro("Insert several blocks")}
  #   adds.jsonl:
  #     {"after":"p3","text":"New clause.","style":"Heading2"}
  #     {"before":"p0","text":"ALERT","color":"CC0000","bold":true}
  #     {"after":"p5","markdown":"## Summary"}
  docx insert doc.docx --batch adds.jsonl
  # …or one at a time:
  docx insert doc.docx --after p3 --text "Section header" --style Heading2
  docx insert doc.docx --after p3 --text "click here" --url https://example.com
  docx insert doc.docx --after p3 --page-break
  docx insert doc.docx --after p3 --markdown "## New section"
  docx insert doc.docx --at-start --text "Title" --style Title
  docx insert doc.docx --after p3 --text-file reviewer-notes.txt

Ordering: batch entries apply in file order; several anchored after the SAME
block stack in that order (three "after":"p0" land as p1, p2, p3, not reversed).

Placement (exactly one required) — where to put the new block:
  --after LOCATOR   Insert after the block at LOCATOR
  --before LOCATOR  Insert before the block at LOCATOR
                    LOCATOR is one of:
${ANCHOR_FORMS}
  --at-start        Insert at the very top (before the first block) — no locator.
  --at-end          Insert at the very end (after the last block, before the
                    trailing section properties) — no locator.
                    (--at-start/--at-end are single-shot only, not --batch.)

Content (one required):
  --markdown TEXT   Parse TEXT as GFM markdown → one or more blocks (headings,
                    lists, tables, code fences, blockquotes, rules, links, inline
                    + display math, images, footnotes, ~~strike~~, CriticMarkup).
  --markdown-file PATH  Same as --markdown, but read from PATH ("-" = stdin).
  --text TEXT       Insert a paragraph with this text (one run). Format it with
                    --bold/--italic/--color/--url — see \`docx insert --text --help\`.
  --text-file PATH  Insert literal multi-paragraph text from PATH ("-" = stdin),
                    NOT parsed as markdown — every character verbatim, each newline
                    a new paragraph. Use for prose that must stay untouched ("3.
                    note" stays "3.", bare URLs / *x* / {++x++} not interpreted).
  --runs JSON       Insert a paragraph with custom runs (Run[] JSON).
                    See \`docx insert --runs --help\`.
  --page-break      Insert an empty paragraph containing a page break
  --column-break    Insert an empty paragraph containing a column break

Formatting options (incompatible with --markdown / --markdown-file):
  --style NAME       Apply paragraph style (e.g., Heading1)
  --alignment ALIGN  left | center | right | justify
  --space-before PT / --space-after PT   Space above / below, in points
  --line-spacing N   A multiple (1, 1.5, 2), a name, or 15pt
  --indent-left IN / --indent-right IN   Indent, in inches
  --first-line IN / --hanging IN         First-line / hanging indent, in inches
  --list KIND        Make the paragraph a list item: "bullet" or "ordered"
                     (requires --text/--runs; task checkbox → \`docx tasks add\`).
  --list-level N     List nesting level, integer 0-8 (use with --list to nest).

Batch (--batch PATH | -):
  Apply many inserts from one read — the preferred way to add several blocks
  (locators do not shift between entries). Each JSONL line is one insert whose
  keys mirror the flags: {"after" or "before": LOCATOR, one content field,
  ...options}, e.g. {"after":"p3","text":"Hi","style":"Heading2"}.
  (--at-start/--at-end don't work in a batch.) Don't pass --after/--text/…
  alongside --batch.

General options:
  --author NAME     Author for tracked changes (default: $DOCX_AUTHOR)
  --track           Record this insertion as a tracked change even when the
                    document's track-changes toggle is off (OFF by default).
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would be inserted; do not write the file
  -v, --verbose     Print the full success ack JSON
  -h, --help        Show this help

Output:
  Prints the locator(s) the new block(s) landed at, one per line (a multi-block
  --markdown insert prints several). Positional ids shift after an insert, so
  re-read before further edits. --verbose prints {ok:true, operation, path,
  locators, anchor, placement}. Errors print {code, error, hint?} + nonzero exit.
`;

const INSERT_TEXT_HELP = `docx insert --text — insert new text content and format it

Usage:
  docx insert FILE (--after | --before) LOCATOR --text "New paragraph" [options]
  docx insert FILE --at-start --text "Title" --style Title

Examples:
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
	const placement = await parseTargetPlacement(values, INSERT_HELP);
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

	// Markdown spec carries its own block styling (heading levels, list
	// numbering, code blocks, …) so paragraph-level flags would be silently
	// dropped. Reject them up front instead.
	if (spec.kind === "markdown") {
		const conflict = MARKDOWN_INCOMPATIBLE_FLAGS.find(
			(flag) => values[flag] !== undefined,
		);
		if (conflict) {
			return fail(
				"USAGE",
				`--${conflict} can't be combined with --markdown / --markdown-file (the markdown source controls block-level styling)`,
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
	};
}

const OPTION_SPEC = {
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
};

export type RawValues = ReturnType<typeof parseArgs>["values"];

/** The mutually-exclusive content flags, each with the sub-flags that only
 * make sense alongside it. Drives both the "exactly one content flag" check
 * and the "this sub-flag requires its content flag" check, so those rules
 * live in one place instead of scattered guards. */
/** Paragraph-level flags that are meaningless under `--markdown` /
 *  `--markdown-file` because the markdown source already encodes block
 *  styling (heading levels, list numbering, code-block fences, …). We
 *  reject explicitly so the agent doesn't silently lose their intent.
 *  `--text` / `--runs` etc. still accept these. */
export const MARKDOWN_INCOMPATIBLE_FLAGS = [
	"style",
	"alignment",
	"list",
	"list-level",
	"space-before",
	"space-after",
	"line-spacing",
	"indent-left",
	"indent-right",
	"first-line",
	"hanging",
] as const;

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
