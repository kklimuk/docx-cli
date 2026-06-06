import {
	type BlockRangeReference,
	type BlockReference,
	type Document,
	describeForms,
	Edit,
	EditError,
	MarkdownImport,
	MarkdownImportError,
	type ParagraphContentSpec,
	type Run,
	type SectionType,
} from "@core";
import type { ParagraphOptions } from "@core/blocks";
import {
	EquationNotFoundError,
	EquationParseError,
	EquationStaleError,
	Equations,
} from "@core/equation";
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
	resolveBlockRangeOrFail,
	respond,
	respondAck,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const AT_FORMS = describeForms(
	["paragraph", "blockRange", "cellParagraph", "section", "equation"],
	"                      ",
);

const HELP = `docx edit — replace a paragraph (or paragraph range), a section, or an equation

Usage:
  docx edit FILE --at LOCATOR <content> [options]

Locator (required):
  --at LOCATOR      What to edit. One of:
${AT_FORMS}
                    pN / pN-pM take paragraph content (below); sN takes
                    --columns/--type; eqN takes --equation/--display/--inline.
                    See \`docx info locators\`.

Paragraph content (one required for paragraph / range locators):
  --text TEXT       Replace with a single-run paragraph
  --runs JSON       Replace with custom runs (Run[] JSON)
  --code TEXT       Replace with a code block — newlines split into one
                    CodeBlock-styled paragraph per source line
  --code-file PATH  Same as --code, but read content from PATH (use "-" for stdin)
  --markdown TEXT   Replace with parsed GFM markdown. Same dialect as
                    'docx insert --markdown' (headings, lists, tables, code,
                    blockquotes, links, math, footnotes, CriticMarkup, …).
                    Multi-block sources expand naturally — a paragraph
                    locator gets replaced by however many blocks the source
                    parses to.
  --markdown-file PATH  Same as --markdown, but read content from PATH
                    (use "-" for stdin).
  --task STATE      Flip a task-list item's checkbox state in place ("checked" or
                    "unchecked"). Requires a single paragraph locator that is
                    already a GFM task list item (has a leading <w14:checkbox/>
                    SDT). Under track-changes, emits Word's canonical toggle
                    shape (ins/del pair inside sdtContent + w14:checked flip);
                    "track-changes list" surfaces it as a checkboxToggle revision.

Equation editing (requires --at eqN):
  --equation LATEX  Replace the equation's content with new LaTeX. Goes through
                    temml → MathML → OMML.
  --display         Switch the equation to display mode (block, $$…$$). Can be
                    combined with --equation, or used alone to toggle mode.
  --inline          Switch to inline mode ($…$). Mutex with --display.

Paragraph options:
  --style NAME       Paragraph style (e.g., Heading1)
  --alignment ALIGN  left | center | right | justify

Run options (only with --text):
  --color HEX       Run color, hex (e.g., 800080 for purple)
  --bold            Bold
  --italic          Italic

Code options (only with --code / --code-file):
  --language LANG   Syntax-highlight via lowlight (37 common languages bundled).
                    Survives round-trip via a CodeBlock-LANG pStyle suffix.

Section options (for section locators sN):
  --columns N        Number of columns for the targeted section
  --type T           continuous | nextPage | evenPage | oddPage | nextColumn

Formatting (single-paragraph --text only):
  By default --text preserves run-level formatting (bold/italic/color/etc.)
  on words shared between the old and new text via a word-level diff. New
  words inherit formatting from the nearest unchanged neighbor. Pass
  --no-formatting to fall back to a single fresh run with no formatting.
  Passing --color/--bold/--italic also bypasses preservation — those flags
  apply uniformly to the new paragraph. Range edits (pN-pM) always rewrite
  the span wholesale; per-word formatting preservation is not applied.

General options:
  --author NAME     Author for tracked changes (default: $DOCX_AUTHOR)
  --no-formatting   Replace with a single fresh run; do not preserve rPr
                    on unchanged words
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -v, --verbose     Print the success ack JSON (default: silent on success)
  -h, --help        Show this help

Output:
  Silent on success (exit 0) — the edited locator is unchanged, so there's
  nothing to mint. --verbose prints {ok:true, operation, path, locator}.
  Errors print {code, error, hint?} with a nonzero exit. Discover ids with
  \`docx read FILE --ast\` (equation ids appear on EquationRun nodes).

Examples:
  docx edit doc.docx --at p3 --text "Replaced." --style Heading2
  docx edit doc.docx --at p0 --runs '[{"type":"text","text":"X","bold":true}]'
  docx edit doc.docx --at p2-p5 --text "Rewrite this section as one paragraph."
  docx edit doc.docx --at p3-p7 --code-file new-snippet.go --language go
  docx edit doc.docx --at s0 --columns 2 --type continuous
  docx edit doc.docx --at eq0 --equation "x = \\\\frac{-b}{2a}" --display
`;

export async function run(args: string[]): Promise<number> {
	const opts = await parseAndValidateOptions(args);
	if (typeof opts === "number") return opts;

	const document = await openOrFail(opts.filePath);
	if (typeof document === "number") return document;

	// Range locator (`pN-pM`): replaces a span of paragraphs as a unit. Section
	// edits don't make sense here (sN has its own grammar).
	if (isBlockRangeLocator(opts.locator)) {
		if (opts.spec.kind === "section") {
			return fail(
				"USAGE",
				"Range locators (pN-pM) don't accept --columns/--type — use sN for section edits",
				HELP,
			);
		}
		return commitRangeEdit(document, opts);
	}

	// Equation locator (`eqN`) targets an `<m:oMath>` inside a paragraph, not
	// the paragraph itself — resolves via the Equations lens, not the block
	// resolver.
	if (opts.spec.kind === "equation") {
		return commitEquationEdit(document, opts.spec, opts);
	}

	const blockRef = await resolveBlockOrFail(document, opts.locator);
	if (typeof blockRef === "number") return blockRef;

	return commitBlockEdit(document, blockRef, opts);
}

function isBlockRangeLocator(locator: string): boolean {
	return /^p\d+-p\d+$/.test(locator);
}

/** Single-block edit: section / task / paragraph dispatch through the Edit
 * lens. The lens handles tracked-vs-untracked, style ensures, and the
 * formatting-preservation decision. */
async function commitBlockEdit(
	document: Document,
	blockRef: BlockReference,
	opts: ValidatedOptions,
): Promise<number> {
	if (opts.dryRun) return respondDryRun(opts);

	try {
		const edit = new Edit(document);
		if (opts.spec.kind === "section") {
			edit.section(blockRef, opts.spec, { authorFlag: opts.authorFlag });
		} else if (opts.spec.kind === "task") {
			edit.taskToggle(blockRef, opts.spec.checked, {
				authorFlag: opts.authorFlag,
			});
		} else if (
			opts.spec.kind === "text" ||
			opts.spec.kind === "runs" ||
			opts.spec.kind === "code"
		) {
			edit.paragraph(blockRef, opts.spec, {
				authorFlag: opts.authorFlag,
				noFormatting: opts.noFormatting,
			});
		} else if (opts.spec.kind === "markdown") {
			const resolved = await resolveMarkdownBlocks(document, opts.spec);
			if (typeof resolved === "number") return resolved;
			edit.paragraph(blockRef, resolved, { authorFlag: opts.authorFlag });
		} else {
			return fail("USAGE", "Unsupported edit spec for single-block locator");
		}
	} catch (error) {
		if (error instanceof EditError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	await document.save(opts.outputPath);
	return emitEditAck(opts);
}

/** Range replace path (`pN-pM`): resolve to a block range, hand to the Edit
 *  lens. The lens rejects tracked ranges that span a non-paragraph block. */
async function commitRangeEdit(
	document: Document,
	opts: ValidatedOptions,
): Promise<number> {
	if (opts.spec.kind === "section") {
		// Type narrowing — caller already rejected this above.
		return fail("USAGE", "Section edits don't support range locators", HELP);
	}
	if (opts.spec.kind === "task") {
		return fail(
			"USAGE",
			"--task takes a single paragraph locator (pN), not a range",
			HELP,
		);
	}
	if (opts.spec.kind === "equation") {
		return fail(
			"USAGE",
			"--equation takes a single equation locator (eqN), not a paragraph range",
			HELP,
		);
	}

	const rangeRef: BlockRangeReference | number = await resolveBlockRangeOrFail(
		document,
		opts.locator,
	);
	if (typeof rangeRef === "number") return rangeRef;

	if (opts.dryRun) return respondDryRun(opts);

	let spec: ParagraphContentSpec;
	if (opts.spec.kind === "markdown") {
		const resolved = await resolveMarkdownBlocks(document, opts.spec);
		if (typeof resolved === "number") return resolved;
		spec = resolved;
	} else if (
		opts.spec.kind === "text" ||
		opts.spec.kind === "runs" ||
		opts.spec.kind === "code"
	) {
		spec = opts.spec;
	} else {
		return fail("USAGE", "Unsupported edit spec for range locator");
	}

	try {
		new Edit(document).range(rangeRef, spec, {
			authorFlag: opts.authorFlag,
		});
	} catch (error) {
		if (error instanceof EditError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	await document.save(opts.outputPath);
	return emitEditAck(opts);
}

/** Resolve an `eqN` locator, splice in a new OMML subtree, save. The spec
 *  carries optional `latex` (content swap) and `display` (mode toggle).
 *  At least one must change something, else it's a no-op error.
 *
 *  Tracking: when `<w:trackChanges/>` is on, the splice is replaced by a
 *  paired `<w:del>OLD</w:del><w:ins>NEW</w:ins>`. Word's accept-all picks
 *  the right equation but leaves an empty container skeleton next to it —
 *  a Word normalization quirk that's cosmetic. */
async function commitEquationEdit(
	document: Document,
	spec: Extract<EditSpec, { kind: "equation" }>,
	opts: ValidatedOptions,
): Promise<number> {
	if (spec.latex === undefined && spec.display === undefined) {
		return fail(
			"USAGE",
			"--equation requires --equation NEW_LATEX, --display, or --inline",
		);
	}

	if (opts.dryRun) return respondDryRun(opts);

	try {
		new Equations(document).edit(spec.locator, {
			latex: spec.latex,
			display: spec.display,
			author: opts.authorFlag,
		});
	} catch (error) {
		if (error instanceof EquationNotFoundError) {
			return fail("BLOCK_NOT_FOUND", error.message);
		}
		if (error instanceof EquationStaleError) {
			return fail("BLOCK_NOT_FOUND", error.message);
		}
		if (error instanceof EquationParseError) {
			return fail(
				"USAGE",
				`Could not parse LaTeX equation: ${error.message}`,
				"Check the LaTeX syntax — temml accepts most KaTeX/MathJax LaTeX.",
			);
		}
		throw error;
	}

	await document.save(opts.outputPath);
	return emitEditAck(opts);
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

	const locator = parsed.values.at as string | undefined;
	if (!locator) return fail("USAGE", "Missing --at LOCATOR", HELP);

	const paragraphOptions = await parseParagraphOptions(parsed.values);
	if (typeof paragraphOptions === "number") return paragraphOptions;

	const isSectionLocator = /^s\d+$/.test(locator);
	const spec = isSectionLocator
		? await validateSectionEdit(parsed.values)
		: await validateParagraphEdit(parsed.values, paragraphOptions);
	if (typeof spec === "number") return spec;

	return {
		filePath,
		locator,
		spec,
		authorFlag: parsed.values.author as string | undefined,
		outputPath: parsed.values.output as string | undefined,
		dryRun: Boolean(parsed.values["dry-run"]),
		noFormatting: Boolean(parsed.values["no-formatting"]),
	};
}

/** Paragraph-level flags that are meaningless under `--markdown` /
 *  `--markdown-file` (the markdown source already encodes block styling).
 *  See `chooseContentSpec` in `cli/insert/index.tsx` for the symmetric
 *  rejection on the insert side. */
const MARKDOWN_INCOMPATIBLE_FLAGS = ["style", "alignment"] as const;

const OPTION_SPEC = {
	at: { type: "string" },
	text: { type: "string" },
	runs: { type: "string" },
	code: { type: "string" },
	"code-file": { type: "string" },
	markdown: { type: "string" },
	"markdown-file": { type: "string" },
	language: { type: "string" },
	task: { type: "string" },
	equation: { type: "string" },
	display: { type: "boolean" },
	inline: { type: "boolean" },
	columns: { type: "string" },
	type: { type: "string" },
	style: { type: "string" },
	alignment: { type: "string" },
	color: { type: "string" },
	bold: { type: "boolean" },
	italic: { type: "boolean" },
	author: { type: "string" },
	"no-formatting": { type: "boolean" },
	output: { type: "string", short: "o" },
	"dry-run": { type: "boolean" },
	verbose: { type: "boolean", short: "v" },
	help: { type: "boolean", short: "h" },
} as const;

type ValidatedOptions = {
	filePath: string;
	locator: string;
	spec: EditSpec;
	authorFlag?: string;
	outputPath?: string;
	dryRun: boolean;
	/** Opt-out of word-level formatting preservation. When true, --text
	 *  produces a single fresh `<w:r>` with no rPr (today's behavior). */
	noFormatting: boolean;
};

type EditSpec =
	| { kind: "section"; columns?: number; sectionType?: SectionType }
	| {
			kind: "text";
			text: string;
			format: TextFormatting;
			paragraphOptions: ParagraphOptions;
	  }
	| { kind: "runs"; runs: Run[]; paragraphOptions: ParagraphOptions }
	| {
			kind: "code";
			content: string;
			language?: string;
			paragraphOptions: ParagraphOptions;
	  }
	| {
			kind: "markdown";
			source: string;
			paragraphOptions: ParagraphOptions;
	  }
	| { kind: "task"; checked: boolean }
	| {
			kind: "equation";
			locator: string;
			/** `undefined` means "keep the existing LaTeX" (a pure display-mode
			 *  toggle); a string means replace the content. */
			latex: string | undefined;
			/** `undefined` means "keep the existing display flag"; a boolean
			 *  switches to that mode. */
			display: boolean | undefined;
	  };

type TextFormatting = {
	color?: string;
	bold?: boolean;
	italic?: boolean;
};

type RawValues = ReturnType<typeof parseArgs>["values"];

async function validateSectionEdit(
	values: RawValues,
): Promise<EditSpec | number> {
	if (values.text !== undefined || values.runs !== undefined) {
		return fail(
			"USAGE",
			"Section locators (sN) take --columns and --type, not --text/--runs",
			HELP,
		);
	}
	if (values.columns === undefined && values.type === undefined) {
		return fail("USAGE", "Section edit requires --columns and/or --type", HELP);
	}
	const sectionFlags = await parseSectionFlags(values);
	if (typeof sectionFlags === "number") return sectionFlags;
	return { kind: "section", ...sectionFlags };
}

async function validateParagraphEdit(
	values: RawValues,
	paragraphOptions: ParagraphOptions,
): Promise<EditSpec | number> {
	if (values.columns !== undefined || values.type !== undefined) {
		return fail(
			"USAGE",
			"--columns and --type require a section locator (sN)",
			HELP,
		);
	}
	const text = values.text as string | undefined;
	const runsJson = values.runs as string | undefined;
	const codeInline = values.code as string | undefined;
	const codeFile = values["code-file"] as string | undefined;
	const language = values.language as string | undefined;
	const taskFlag = values.task as string | undefined;
	const equationFlag = values.equation as string | undefined;
	const displayFlag = values.display as boolean | undefined;
	const inlineFlag = values.inline as boolean | undefined;

	// `--equation` swaps the LaTeX (and optionally display mode) of an
	// `eqN`-addressed equation. `--display` / `--inline` alone (without
	// `--equation`) toggle the display mode but keep the existing LaTeX.
	const equationLike =
		equationFlag !== undefined || displayFlag === true || inlineFlag === true;
	if (equationLike) {
		const otherContent = [
			text !== undefined,
			runsJson !== undefined,
			codeInline !== undefined,
			codeFile !== undefined,
			taskFlag !== undefined,
		].filter(Boolean).length;
		if (otherContent > 0) {
			return fail(
				"USAGE",
				"--equation / --display / --inline cannot be combined with --text/--runs/--code/--code-file/--task",
				HELP,
			);
		}
		if (displayFlag && inlineFlag) {
			return fail(
				"USAGE",
				"--display and --inline are mutually exclusive",
				HELP,
			);
		}
		const displayMode: boolean | undefined = displayFlag
			? true
			: inlineFlag
				? false
				: undefined;
		return {
			kind: "equation",
			locator: values.at as string,
			latex: equationFlag,
			display: displayMode,
		};
	}

	// `--task` is its own content kind — it flips an existing task's state in
	// place rather than replacing the paragraph.
	if (taskFlag !== undefined) {
		const otherFlags = [
			text !== undefined,
			runsJson !== undefined,
			codeInline !== undefined,
			codeFile !== undefined,
		].filter(Boolean).length;
		if (otherFlags > 0) {
			return fail(
				"USAGE",
				"--task cannot be combined with --text, --runs, --code, or --code-file",
				HELP,
			);
		}
		const checked = parseTaskFlag(taskFlag);
		if (checked === null) {
			return fail(
				"USAGE",
				`--task must be "checked" or "unchecked", got "${taskFlag}"`,
				HELP,
			);
		}
		return { kind: "task", checked };
	}

	const markdownInline = values.markdown as string | undefined;
	const markdownFile = values["markdown-file"] as string | undefined;
	const contentFlags = [
		text !== undefined,
		runsJson !== undefined,
		codeInline !== undefined,
		codeFile !== undefined,
		markdownInline !== undefined,
		markdownFile !== undefined,
	].filter(Boolean).length;
	if (contentFlags === 0) {
		return fail(
			"USAGE",
			"Missing content: pass --text, --runs, --code, --code-file, --markdown, --markdown-file, --task, or --equation",
			HELP,
		);
	}
	if (contentFlags > 1) {
		return fail(
			"USAGE",
			"Pass only one of --text, --runs, --code, --code-file, --markdown, --markdown-file",
			HELP,
		);
	}
	if (
		language !== undefined &&
		codeInline === undefined &&
		codeFile === undefined
	) {
		return fail("USAGE", "--language requires --code or --code-file", HELP);
	}

	if (text !== undefined) {
		return {
			kind: "text",
			text,
			format: {
				color: values.color as string | undefined,
				bold: values.bold as boolean | undefined,
				italic: values.italic as boolean | undefined,
			},
			paragraphOptions,
		};
	}

	if (codeInline !== undefined || codeFile !== undefined) {
		const content =
			codeInline !== undefined
				? codeInline
				: await loadCodeFile(codeFile as string);
		if (typeof content === "number") return content;
		return {
			kind: "code",
			content,
			...(language ? { language } : {}),
			paragraphOptions,
		};
	}

	if (markdownInline !== undefined || markdownFile !== undefined) {
		// Markdown encodes its own block styling — paragraph-level flags
		// would be silently dropped. Reject up front for the same reason
		// `--language` is gated to `--code` / `--code-file`.
		const conflict = MARKDOWN_INCOMPATIBLE_FLAGS.find(
			(flag) => values[flag] !== undefined,
		);
		if (conflict) {
			return fail(
				"USAGE",
				`--${conflict} can't be combined with --markdown / --markdown-file (the markdown source controls block-level styling)`,
				HELP,
			);
		}
		const source =
			markdownInline !== undefined
				? markdownInline
				: await loadMarkdownFile(markdownFile as string);
		if (typeof source === "number") return source;
		return { kind: "markdown", source, paragraphOptions };
	}

	const runs = await parseRunsArg(runsJson as string);
	if (typeof runs === "number") return runs;
	return { kind: "runs", runs, paragraphOptions };
}

/** Read content for `--markdown-file PATH`. `-` means stdin. */
async function loadMarkdownFile(path: string): Promise<string | number> {
	try {
		return path === "-"
			? await new Response(Bun.stdin.stream()).text()
			: await Bun.file(path).text();
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail(
			"FILE_NOT_FOUND",
			`Failed to read --markdown-file ${path}: ${message}`,
		);
	}
}

/** Resolve a `markdown` spec into a `markdown-blocks` spec by parsing the
 * source against the now-open document. The lens registers footnote bodies,
 * mints image rels, and provisions any styles the source references — all on
 * `document` — before returning the splice-ready blocks. */
async function resolveMarkdownBlocks(
	document: Document,
	spec: Extract<EditSpec, { kind: "markdown" }>,
): Promise<
	Extract<ParagraphContentSpec, { kind: "markdown-blocks" }> | number
> {
	try {
		const blocks = await new MarkdownImport(document).blocks(spec.source);
		return {
			kind: "markdown-blocks",
			blocks,
			paragraphOptions: spec.paragraphOptions,
		};
	} catch (error) {
		if (error instanceof MarkdownImportError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}
}

/** Read content for `--code-file PATH`. `-` means stdin (handled the same as
 *  `insert --code-file -`). Mirrors `cli/insert/index.tsx::resolveCodeSpec`. */
async function loadCodeFile(path: string): Promise<string | number> {
	try {
		return path === "-"
			? await new Response(Bun.stdin.stream()).text()
			: await Bun.file(path).text();
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail(
			"FILE_NOT_FOUND",
			`Failed to read --code-file ${path}: ${message}`,
		);
	}
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

	return out;
}

async function respondDryRun(opts: ValidatedOptions): Promise<number> {
	await respond({
		operation: "edit",
		dryRun: true,
		path: opts.filePath,
		locator: opts.locator,
		...(opts.outputPath ? { output: opts.outputPath } : {}),
	});
	return EXIT.OK;
}

async function emitEditAck(opts: ValidatedOptions): Promise<number> {
	await respondAck({
		ok: true,
		operation: "edit",
		path: opts.outputPath ?? opts.filePath,
		locator: opts.locator,
	});
	return EXIT.OK;
}
