import {
	type BlockReference,
	type Document,
	Insert,
	InsertError,
	type InsertSpec,
} from "@core";
import type { ParagraphOptions } from "@core/blocks";
import type { XmlNode } from "@core/parser";
import type { TableBorders, TableLayout } from "@core/table";
import type { parseArgs } from "util";
import {
	parseRunsArg,
	parseSectionFlags,
	parseTaskFlag,
} from "../parse-helpers";
import {
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
	respond,
	respondAck,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx insert — insert a paragraph at a locator

Usage:
  docx insert FILE [options]

Locator (one required):
  --after LOCATOR   Insert after the block at LOCATOR (e.g., p3)
  --before LOCATOR  Insert before the block at LOCATOR

Content (one required):
  --text TEXT       Insert a paragraph with this text
  --runs JSON       Insert a paragraph with custom runs (Run[] JSON)
  --page-break      Insert an empty paragraph containing a page break
  --column-break    Insert an empty paragraph containing a column break
  --section         Insert a section boundary (sentinel paragraph w/ inline sectPr)
  --table           Insert an empty rows×cols table (requires --rows and --cols)
  --image SRC       Insert an image (SRC is a file path, data: URI, or http(s) URL)
  --code TEXT       Insert a multi-line code block. Newlines split into one
                    CodeBlock-styled paragraph per source line.
  --code-file PATH  Same as --code, but read content from PATH (use "-" for stdin).
  --equation LATEX  Insert a math equation from LaTeX. Goes through temml
                    (KaTeX/MathJax-compatible LaTeX dialect) → MathML → OMML.
                    Pair with --display for block-mode equations; omit for
                    inline. Round-trips as $LATEX$ / $$LATEX$$ in markdown.
  --markdown TEXT   Parse TEXT as GFM markdown (remark + remark-gfm +
                    remark-math + CriticMarkup) and emit the result as one or
                    more blocks. Supports: headings, paragraphs, lists
                    (bullet/ordered/task), tables, code fences, blockquotes,
                    horizontal rules, links, inline + display math ($x^2$,
                    $$x^2$$), inline images (path/URL/data:), footnote
                    refs/defs ([^id]: body), GFM strikethrough (~~x~~), and
                    CriticMarkup insertions/deletions ({++x++}/{--x--}; under
                    tracking these wrap in <w:ins>/<w:del>, otherwise the
                    insertion is plain text and the deletion is dropped).
  --markdown-file PATH  Same as --markdown, but read content from PATH
                    (use "-" for stdin).

Paragraph options:
  --style NAME       Apply paragraph style (e.g., Heading1)
  --alignment ALIGN  left | center | right | justify
  --task STATE       Make the new paragraph a GFM task list item with state
                     "checked" (☒) or "unchecked" (☐). Requires --text or --runs.
                     If the anchor is itself a list paragraph, inherits its numId
                     (so consecutive --task inserts build a contiguous list);
                     otherwise allocates a fresh bullet list.

Run options (only with --text):
  --color HEX       Run color, hex (e.g., 800080 for purple)
  --bold            Bold
  --italic          Italic
  --url URL         Wrap the inserted text in a hyperlink to URL

Section options (only with --section):
  --columns N       Number of columns for the section ending at this boundary
  --type T          continuous | nextPage | evenPage | oddPage | nextColumn

Table options (only with --table):
  --rows N          Number of rows (required, >= 1)
  --cols N          Number of columns (required, >= 1)
  --widths "A,B,C"  Column widths in twips, comma-separated; length must equal --cols
  --table-width V   Table total width, e.g. "100%" (default), "50%", or "4320" (twips)
  --borders S       single (default) | none | double
  --layout L        autofit (default; columns size to content) | fixed (honor
                    --widths exactly). Passing --widths implies fixed.

Image options (only with --image):
  --alt TEXT        Alt text / description for the image
  --width INCHES    Display width in inches (default: native pixel size at 96dpi)
  --height INCHES   Display height in inches (default: scales to preserve aspect)

Code options (only with --code / --code-file):
  --language LANG   Syntax-highlight using lowlight (highlight.js). One of the
                    37 common languages: bash, c, cpp, csharp, css, diff, go,
                    graphql, ini, java, javascript, json, kotlin, less, lua,
                    makefile, markdown, objectivec, perl, php, php-template,
                    plaintext, python, python-repl, r, ruby, rust, scss, shell,
                    sql, swift, typescript, vbnet, wasm, xml, yaml. Unknown
                    languages degrade to uncolored (block still inserts).

General options:
  --author NAME     Author for tracked changes (default: $DOCX_AUTHOR)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would be inserted; do not write the file
  -v, --verbose     Print the success ack JSON (default: silent on success)
  -h, --help        Show this help

Examples:
  docx insert doc.docx --after p3 --text "Section header" --style Heading2
  docx insert doc.docx --before p0 --text "ALERT" --color CC0000 --bold
  docx insert doc.docx --after p2 --runs '[{"type":"text","text":"X","bold":true}]'
  docx insert doc.docx --after p3 --text "click here" --url https://example.com
  docx insert doc.docx --after p3 --page-break
  docx insert doc.docx --after p9 --section --columns 2 --type continuous
  docx insert doc.docx --after p3 --table --rows 3 --cols 2
  docx insert doc.docx --after p3 --table --rows 2 --cols 3 --widths 1440,2880,4320
  docx insert doc.docx --after p3 --image ./diagram.png --alt "System diagram"
  docx insert doc.docx --after p3 --image https://example.com/logo.png --width 2
  docx insert doc.docx --after p3 --code $'function foo() {\\n  return 42;\\n}' --language typescript
  docx insert doc.docx --after p3 --code-file snippet.go --language go
  cat snippet.py | docx insert doc.docx --after p3 --code-file - --language python
  docx insert doc.docx --after p3 --markdown $'# Heading\\n\\n- a\\n- b'
  docx insert doc.docx --after p3 --markdown-file README.md
  cat draft.md | docx insert doc.docx --after p3 --markdown-file -
`;

export async function run(args: string[]): Promise<number> {
	const opts = await parseAndValidateOptions(args);
	if (typeof opts === "number") return opts;

	const document = await openOrFail(opts.filePath);
	if (typeof document === "number") return document;

	const blockRef = await resolveBlockOrFail(document, opts.placement.locator);
	if (typeof blockRef === "number") return blockRef;

	let blocks: Awaited<ReturnType<Insert["paragraph"]>>;
	try {
		blocks = await new Insert(document).paragraph(
			blockRef,
			opts.spec,
			opts.paragraphOptions,
			{ placement: opts.placement.mode, authorFlag: opts.authorFlag },
		);
	} catch (error) {
		if (error instanceof InsertError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	return commitInsert(document, blockRef, blocks, opts);
}

/** Splice the built blocks into the document and persist (unless `--dry-run`).
 * Kept as a small CLI helper so the response/output-path orchestration stays
 * next to `run()`. */
async function commitInsert(
	document: Document,
	blockRef: BlockReference,
	blocks: XmlNode[],
	opts: ValidatedOptions,
): Promise<number> {
	if (opts.dryRun) {
		await respond({
			ok: true,
			operation: "insert",
			dryRun: true,
			path: opts.filePath,
			locator: opts.placement.locator,
			placement: opts.placement.mode,
			...(opts.outputPath ? { output: opts.outputPath } : {}),
		});
		return EXIT.OK;
	}

	const targetIndex = blockRef.parent.indexOf(blockRef.node);
	if (targetIndex === -1) {
		return fail(
			"BLOCK_NOT_FOUND",
			"Block reference is stale (parent does not contain it)",
		);
	}
	const insertIndex =
		opts.placement.mode === "after" ? targetIndex + 1 : targetIndex;
	blockRef.parent.splice(insertIndex, 0, ...blocks);
	await document.save(opts.outputPath);

	await respondAck({
		ok: true,
		operation: "insert",
		path: opts.outputPath ?? opts.filePath,
		locator: opts.placement.locator,
		placement: opts.placement.mode,
	});
	return EXIT.OK;
}

async function parseAndValidateOptions(
	args: string[],
): Promise<ValidatedOptions | number> {
	const parsed = await tryParseArgs(args, OPTION_SPEC, HELP);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const filePath = parsed.positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE argument", HELP);

	const placement = await parseTargetPlacement(parsed.values);
	if (typeof placement === "number") return placement;

	const spec = await chooseContentSpec(parsed.values);
	if (typeof spec === "number") return spec;

	// Markdown spec carries its own block styling (heading levels, list
	// numbering, code blocks, …) so paragraph-level flags would be silently
	// dropped. Reject them up front instead.
	if (spec.kind === "markdown") {
		const conflict = MARKDOWN_INCOMPATIBLE_FLAGS.find(
			(flag) => parsed.values[flag] !== undefined,
		);
		if (conflict) {
			return fail(
				"USAGE",
				`--${conflict} can't be combined with --markdown / --markdown-file (the markdown source controls block-level styling)`,
				HELP,
			);
		}
	}

	const paragraphOptions = await parseParagraphOptions(parsed.values);
	if (typeof paragraphOptions === "number") return paragraphOptions;

	return {
		filePath,
		placement,
		spec,
		paragraphOptions,
		authorFlag: parsed.values.author as string | undefined,
		outputPath: parsed.values.output as string | undefined,
		dryRun: Boolean(parsed.values["dry-run"]),
	};
}

const OPTION_SPEC = {
	after: { type: "string" },
	before: { type: "string" },
	text: { type: "string" },
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
	color: { type: "string" },
	bold: { type: "boolean" },
	italic: { type: "boolean" },
	url: { type: "string" },
	author: { type: "string" },
	output: { type: "string", short: "o" },
	"dry-run": { type: "boolean" },
	verbose: { type: "boolean", short: "v" },
	help: { type: "boolean", short: "h" },
} as const;

type ValidatedOptions = {
	filePath: string;
	placement: { mode: "after" | "before"; locator: string };
	spec: InsertSpec;
	paragraphOptions: ParagraphOptions;
	authorFlag?: string;
	outputPath?: string;
	dryRun: boolean;
};

async function parseTargetPlacement(
	values: RawValues,
): Promise<{ mode: "after" | "before"; locator: string } | number> {
	const after = values.after as string | undefined;
	const before = values.before as string | undefined;
	if (!after && !before) {
		return fail("USAGE", "Missing locator: pass --after or --before", HELP);
	}
	if (after && before) {
		return fail("USAGE", "Pass either --after or --before, not both", HELP);
	}
	if (after !== undefined) return { mode: "after", locator: after };
	return { mode: "before", locator: before as string };
}

type RawValues = ReturnType<typeof parseArgs>["values"];

/** The mutually-exclusive content flags, each with the sub-flags that only
 * make sense alongside it. Drives both the "exactly one content flag" check
 * and the "this sub-flag requires its content flag" check, so those rules
 * live in one place instead of scattered guards. */
/** Paragraph-level flags that are meaningless under `--markdown` /
 *  `--markdown-file` because the markdown source already encodes block
 *  styling (heading levels, list numbering, code-block fences, …). We
 *  reject explicitly so the agent doesn't silently lose their intent.
 *  `--text` / `--runs` etc. still accept these. */
const MARKDOWN_INCOMPATIBLE_FLAGS = [
	"style",
	"alignment",
	"task",
	"list",
	"list-level",
] as const;

const CONTENT_KINDS = [
	{ flag: "text", subFlags: ["color", "bold", "italic", "url"] },
	{ flag: "runs", subFlags: [] },
	{ flag: "page-break", subFlags: [] },
	{ flag: "column-break", subFlags: [] },
	{ flag: "section", subFlags: ["columns", "type"] },
	{
		flag: "table",
		subFlags: ["rows", "cols", "widths", "table-width", "borders", "layout"],
	},
	{ flag: "image", subFlags: ["alt", "width", "height"] },
	{ flag: "code", subFlags: ["language"] },
	{ flag: "code-file", subFlags: ["language"] },
	{ flag: "equation", subFlags: ["display"] },
	{ flag: "markdown", subFlags: [] },
	{ flag: "markdown-file", subFlags: [] },
] as const;

const CONTENT_FLAG_LIST = CONTENT_KINDS.map((kind) => `--${kind.flag}`).join(
	", ",
);

async function chooseContentSpec(
	values: RawValues,
): Promise<InsertSpec | number> {
	const present = CONTENT_KINDS.filter(
		(kind) => values[kind.flag] !== undefined,
	);
	if (present.length > 1) {
		return fail("USAGE", `Pass only one of ${CONTENT_FLAG_LIST}`, HELP);
	}
	const chosen = present[0];
	if (!chosen) {
		return fail("USAGE", `Missing content: pass ${CONTENT_FLAG_LIST}`, HELP);
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
			return fail("USAGE", `--${orphan} requires --${kind.flag}`, HELP);
		}
	}

	switch (chosen.flag) {
		case "text":
			return buildTextSpec(values);
		case "runs": {
			const runs = await parseRunsArg(values.runs as string);
			return typeof runs === "number" ? runs : { kind: "runs", runs };
		}
		case "page-break":
			return { kind: "break", breakKind: "page" };
		case "column-break":
			return { kind: "break", breakKind: "column" };
		case "section": {
			const flags = await parseSectionFlags(values);
			return typeof flags === "number" ? flags : { kind: "section", ...flags };
		}
		case "table": {
			const flags = await parseTableFlags(values);
			return typeof flags === "number" ? flags : { kind: "table", ...flags };
		}
		case "image": {
			const flags = await parseImageFlags(values);
			return typeof flags === "number" ? flags : { kind: "image", ...flags };
		}
		case "code":
		case "code-file":
			return resolveCodeSpec(values, chosen.flag);
		case "equation":
			return {
				kind: "equation",
				latex: values.equation as string,
				display: Boolean(values.display),
			};
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
		const source = values.markdown as string;
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

/** Resolve `--code TEXT` (inline) or `--code-file PATH` (file / stdin) into
 *  a uniform `code` spec. Stdin path: `--code-file -` reads from process
 *  stdin so `cat snippet.py | docx insert ... --code-file -` works. */
async function resolveCodeSpec(
	values: RawValues,
	flag: "code" | "code-file",
): Promise<Extract<InsertSpec, { kind: "code" }> | number> {
	const language = values.language as string | undefined;
	if (flag === "code") {
		const content = values.code as string;
		return { kind: "code", content, ...(language ? { language } : {}) };
	}
	const path = values["code-file"] as string;
	try {
		const content =
			path === "-"
				? await new Response(Bun.stdin.stream()).text()
				: await Bun.file(path).text();
		return { kind: "code", content, ...(language ? { language } : {}) };
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail(
			"FILE_NOT_FOUND",
			`Failed to read --code-file ${path}: ${message}`,
		);
	}
}

async function parseImageFlags(values: RawValues): Promise<
	| {
			src: string;
			alt?: string;
			widthInches?: number;
			heightInches?: number;
	  }
	| number
> {
	const src = values.image as string | undefined;
	if (!src) return fail("USAGE", "--image requires a SRC argument", HELP);

	const out: {
		src: string;
		alt?: string;
		widthInches?: number;
		heightInches?: number;
	} = { src };

	const alt = values.alt as string | undefined;
	if (alt !== undefined) out.alt = alt;

	const widthRaw = values.width as string | undefined;
	if (widthRaw !== undefined) {
		const width = Number.parseFloat(widthRaw);
		if (!Number.isFinite(width) || width <= 0) {
			return fail(
				"USAGE",
				`--width must be a positive number of inches, got "${widthRaw}"`,
			);
		}
		out.widthInches = width;
	}

	const heightRaw = values.height as string | undefined;
	if (heightRaw !== undefined) {
		const height = Number.parseFloat(heightRaw);
		if (!Number.isFinite(height) || height <= 0) {
			return fail(
				"USAGE",
				`--height must be a positive number of inches, got "${heightRaw}"`,
			);
		}
		out.heightInches = height;
	}

	return out;
}

function buildTextSpec(
	values: RawValues,
): Extract<InsertSpec, { kind: "text" }> {
	const url = values.url as string | undefined;
	return {
		kind: "text",
		text: values.text as string,
		format: {
			color: values.color as string | undefined,
			bold: values.bold as boolean | undefined,
			italic: values.italic as boolean | undefined,
		},
		...(url ? { hyperlinkUrl: url } : {}),
	};
}

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
		return fail("USAGE", "--table requires --rows and --cols", HELP);
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

async function parseParagraphOptions(
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

	const taskValue = values.task as string | undefined;
	const listValue = values.list as string | undefined;
	const listLevelValue = values["list-level"] as string | undefined;

	if (taskValue !== undefined && listValue !== undefined) {
		return fail(
			"USAGE",
			"--task and --list are mutually exclusive (--task already implies a bullet list)",
			HELP,
		);
	}

	if (taskValue !== undefined) {
		const checked = parseTaskFlag(taskValue);
		if (checked === null) {
			return fail(
				"USAGE",
				`--task must be "checked" or "unchecked", got "${taskValue}"`,
				HELP,
			);
		}
		out.taskState = checked ? "checked" : "unchecked";
	}

	if (listValue !== undefined) {
		if (listValue !== "bullet" && listValue !== "ordered") {
			return fail(
				"USAGE",
				`--list must be "bullet" or "ordered", got "${listValue}"`,
				HELP,
			);
		}
		// Mark the intent to allocate a list; the numId is resolved later in
		// `resolveListContext` (post-document-open) using the same anchor-inherit
		// logic as --task. We stash the kind on a side channel so the resolver
		// knows which abstractNum to use.
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
				HELP,
			);
		}
		if (out.list) out.list.level = level;
		// If neither --task nor --list is set, we still record the level — it
		// applies once the resolver attaches a list (e.g., via inheritance).
		(out as ParagraphOptions & { explicitLevel?: number }).explicitLevel =
			level;
	}

	return out;
}
