import {
	type BlockRangeReference,
	type BlockReference,
	type Document,
	Edit,
	EditError,
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
import { parseArgs } from "util";
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
	writeStdout,
} from "../respond";

const HELP = `docx edit — replace a paragraph (or paragraph range) or modify a section

Usage:
  docx edit FILE [options]

Locator (required):
  --at LOCATOR      What to edit. One of:
                      pN          a single paragraph
                      pN-pM       a contiguous paragraph range (replace as a unit)
                      sN          a section break

Paragraph content (one required for paragraph / range locators):
  --text TEXT       Replace with a single-run paragraph
  --runs JSON       Replace with custom runs (Run[] JSON)
  --code TEXT       Replace with a code block — newlines split into one
                    CodeBlock-styled paragraph per source line
  --code-file PATH  Same as --code, but read content from PATH (use "-" for stdin)
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

Examples:
  docx edit doc.docx --at p3 --text "Replaced." --style Heading2
  docx edit doc.docx --at p0 --runs '[{"type":"text","text":"X","bold":true}]'
  docx edit doc.docx --at p2-p5 --text "Rewrite this section as one paragraph."
  docx edit doc.docx --at p3-p7 --code-file new-snippet.go --language go
  docx edit doc.docx --at s0 --columns 2 --type continuous
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

	try {
		new Edit(document).range(rangeRef, opts.spec, {
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
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: OPTION_SPEC,
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

const OPTION_SPEC = {
	at: { type: "string" },
	text: { type: "string" },
	runs: { type: "string" },
	code: { type: "string" },
	"code-file": { type: "string" },
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

	const contentFlags = [
		text !== undefined,
		runsJson !== undefined,
		codeInline !== undefined,
		codeFile !== undefined,
	].filter(Boolean).length;
	if (contentFlags === 0) {
		return fail(
			"USAGE",
			"Missing content: pass --text, --runs, --code, --code-file, --task, or --equation",
			HELP,
		);
	}
	if (contentFlags > 1) {
		return fail(
			"USAGE",
			"Pass only one of --text, --runs, --code, --code-file",
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

	const runs = await parseRunsArg(runsJson as string);
	if (typeof runs === "number") return runs;
	return { kind: "runs", runs, paragraphOptions };
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
		ok: true,
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
