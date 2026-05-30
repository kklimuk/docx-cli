import {
	applyColumns,
	applySectionType,
	type BlockRangeReference,
	type BlockReference,
	type Document,
	isSectionType,
	LocatorParseError,
	LocatorResolveError,
	parseLocator,
	type Run,
	resolveAuthor,
	resolveBlockRange,
	resolveDate,
	type SectionType,
	TrackChanges,
	type TrackedMeta,
	wrapSectPrChange,
} from "@core";
import { Paragraph, type ParagraphOptions } from "@core/blocks";
import {
	buildCodeBlockParagraphs,
	ensureCodeBlockStyles,
} from "@core/code-block";
import {
	EquationNotFoundError,
	EquationParseError,
	EquationStaleError,
	Equations,
} from "@core/equation";
import type { XmlNode } from "@core/parser";

import { flipCheckboxTracked, flipCheckboxUntracked } from "@core/task-list";
import {
	applyTrackedRangeReplace,
	applyUntrackedRangeReplace,
	applyFormattingPreservingEdit as coreApplyFormattingPreservingEdit,
} from "@core/track-changes/replace";
import { parseArgs } from "util";
import { rejectNonParagraphTrackedRange } from "../range-guard";
import {
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
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

	// Range locator: `pN-pM`. Replaces a span of paragraphs as a unit. Section
	// edits don't make sense here (sN has its own grammar).
	if (isBlockRangeLocator(opts.locator)) {
		if (opts.spec.kind === "section") {
			return fail(
				"USAGE",
				"Range locators (pN-pM) don't accept --columns/--type — use sN for section edits",
				HELP,
			);
		}
		return commitRangeReplacement(document, opts);
	}

	// Equation locator (`eqN`) targets an `<m:oMath>` inside a paragraph,
	// not the paragraph itself — resolve via `equationReferences` instead of
	// the block resolver.
	if (opts.spec.kind === "equation") {
		return commitEquationEdit(document, opts.spec, opts);
	}

	const blockRef = await resolveBlockOrFail(document, opts.locator);
	if (typeof blockRef === "number") return blockRef;

	if (opts.spec.kind === "section") {
		return commitSectionPropertyEdit(document, blockRef, opts.spec, opts);
	}
	if (opts.spec.kind === "task") {
		return commitTaskToggle(document, blockRef, opts.spec, opts);
	}
	document
		.ensureStyles()
		.ensureReferencedStyle(opts.spec.paragraphOptions.style);
	if (opts.spec.kind === "runs") {
		document.ensureStyles().ensureReferencedRunStyles(opts.spec.runs);
	}
	return commitParagraphReplacement(document, blockRef, opts.spec, opts);
}

function isBlockRangeLocator(locator: string): boolean {
	return /^p\d+-p\d+$/.test(locator);
}

/** Range replace path: resolve the blockRange, build new paragraphs from the
 *  spec, splice them in via the tracked or untracked range-replace helper.
 *  No formatting-preservation here — Word's empirical model for paragraph-
 *  range replace is "del all old, ins all new" (no cross-paragraph LCS),
 *  and we match it. */
async function commitRangeReplacement(
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

	let rangeRef: BlockRangeReference;
	try {
		const locator = parseLocator(opts.locator);
		if (locator.kind !== "blockRange") {
			return fail("INVALID_LOCATOR", `Expected pN-pM, got ${opts.locator}`);
		}
		rangeRef = resolveBlockRange(
			document,
			locator.startBlockId,
			locator.endBlockId,
		);
	} catch (err) {
		if (err instanceof LocatorParseError) {
			return fail("INVALID_LOCATOR", err.message);
		}
		if (err instanceof LocatorResolveError) {
			return fail("BLOCK_NOT_FOUND", err.message);
		}
		throw err;
	}

	document
		.ensureStyles()
		.ensureReferencedStyle(opts.spec.paragraphOptions.style);
	if (opts.spec.kind === "runs") {
		document.ensureStyles().ensureReferencedRunStyles(opts.spec.runs);
	}
	if (opts.spec.kind === "code") {
		ensureCodeBlockStyles(document, opts.spec.language);
	}

	const tracked = document.isTrackChangesEnabled();
	if (tracked) {
		const guard = await rejectNonParagraphTrackedRange(rangeRef, opts.locator);
		if (guard !== null) return guard;
	}

	if (opts.dryRun) return respondDryRun(opts);

	const newParagraphs = buildNewParagraphs(opts.spec);
	if (tracked) {
		applyTrackedRangeReplace(
			document,
			rangeRef.parent,
			rangeRef.startIndex,
			rangeRef.endIndex,
			newParagraphs,
			opts.authorFlag,
		);
	} else {
		applyUntrackedRangeReplace(
			rangeRef.parent,
			rangeRef.startIndex,
			rangeRef.endIndex,
			newParagraphs,
		);
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
/** Parse `--task` value into a boolean (checked) or null if unrecognized.
 *  Accepts `checked`/`unchecked` (canonical) plus a few short forms agents
 *  reach for naturally. */
function parseTaskFlag(value: string): boolean | null {
	const normalized = value.toLowerCase();
	if (normalized === "checked" || normalized === "true" || normalized === "1")
		return true;
	if (
		normalized === "unchecked" ||
		normalized === "false" ||
		normalized === "0"
	)
		return false;
	return null;
}

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

async function parseSectionFlags(
	values: RawValues,
): Promise<{ columns?: number; sectionType?: SectionType } | number> {
	const out: { columns?: number; sectionType?: SectionType } = {};

	const columnsRaw = values.columns as string | undefined;
	if (columnsRaw !== undefined) {
		const columns = Number.parseInt(columnsRaw, 10);
		if (!Number.isFinite(columns) || columns <= 0) {
			return fail(
				"USAGE",
				`--columns must be a positive integer, got "${columnsRaw}"`,
			);
		}
		out.columns = columns;
	}

	const sectionTypeRaw = values.type as string | undefined;
	if (sectionTypeRaw !== undefined) {
		if (!isSectionType(sectionTypeRaw)) {
			return fail(
				"USAGE",
				`Invalid --type: ${sectionTypeRaw}`,
				"Valid values: continuous, nextPage, evenPage, oddPage, nextColumn",
			);
		}
		out.sectionType = sectionTypeRaw;
	}

	return out;
}

async function parseRunsArg(json: string): Promise<Run[] | number> {
	let parsed: unknown;
	try {
		parsed = JSON.parse(json);
	} catch (jsonError) {
		const message =
			jsonError instanceof Error ? jsonError.message : String(jsonError);
		return fail("USAGE", `Invalid --runs JSON: ${message}`);
	}
	if (!Array.isArray(parsed)) {
		return fail("USAGE", "--runs must be a JSON array of Run objects");
	}
	return parsed as Run[];
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

async function commitSectionPropertyEdit(
	document: Document,
	blockRef: BlockReference,
	spec: Extract<EditSpec, { kind: "section" }>,
	opts: ValidatedOptions,
): Promise<number> {
	if (blockRef.node.tag !== "w:sectPr") {
		return fail(
			"BLOCK_NOT_FOUND",
			`Locator ${opts.locator} did not resolve to a section break`,
		);
	}

	if (opts.dryRun) return respondDryRun(opts);

	if (document.isTrackChangesEnabled()) {
		wrapSectPrChange(
			blockRef.node,
			new TrackChanges(document).mintMeta(opts.authorFlag),
		);
	}
	applyColumns(blockRef.node, spec.columns);
	applySectionType(blockRef.node, spec.sectionType);

	await document.save(opts.outputPath);
	return emitEditAck(opts);
}

async function commitTaskToggle(
	document: Document,
	blockRef: BlockReference,
	spec: Extract<EditSpec, { kind: "task" }>,
	opts: ValidatedOptions,
): Promise<number> {
	if (blockRef.node.tag !== "w:p") {
		return fail(
			"USAGE",
			"--task requires a paragraph locator; got a non-paragraph block",
		);
	}
	if (opts.dryRun) return respondDryRun(opts);

	const tracked = document.isTrackChangesEnabled();
	let ok: boolean;
	if (tracked) {
		const allocator = new TrackChanges(document).createAllocator();
		const baseMeta = {
			author: resolveAuthor(opts.authorFlag),
			date: resolveDate(),
		};
		const mintMeta = (): TrackedMeta => ({
			...baseMeta,
			revisionId: allocator.next(),
		});
		ok = flipCheckboxTracked(blockRef.node, spec.checked, mintMeta);
	} else {
		ok = flipCheckboxUntracked(blockRef.node, spec.checked);
	}
	if (!ok) {
		return fail(
			"USAGE",
			"--task requires a task-list paragraph (one with a leading <w:sdt><w14:checkbox/></w:sdt>)",
			"Use `docx read FILE --ast` to inspect; convert a plain bullet to a task by replacing the paragraph via `--runs`.",
		);
	}

	await document.save(opts.outputPath);
	return emitEditAck(opts);
}

/** Resolve an `eqN` locator, splice in a new OMML subtree, save. The locator
 *  resolves via `document.body.equationReferences` (populated by the reader); spec
 *  carries optional `latex` (content swap) and `display` (mode toggle). At
 *  least one must change something, else it's a no-op error.
 *
 *  Tracking: when `<w:trackChanges/>` is on, the splice is replaced by a
 *  paired `<w:del>OLD</w:del><w:ins>NEW</w:ins>` pattern next to each other
 *  in the same parent. Our own track-changes accept/reject handles this
 *  cleanly. Word's accept-all also resolves to the correct equation (NEW
 *  on accept, OLD on reject) but leaves an empty `<m:sSup>` / `<m:f>`
 *  structural skeleton next to the kept equation — a Word normalization
 *  quirk that's cosmetic, not a correctness issue (the kept equation
 *  renders right; the skeleton is invisible). */
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

async function commitParagraphReplacement(
	document: Document,
	blockRef: BlockReference,
	spec: ParagraphContentSpec,
	opts: ValidatedOptions,
): Promise<number> {
	const targetIndex = blockRef.parent.indexOf(blockRef.node);
	if (targetIndex === -1) {
		return fail(
			"BLOCK_NOT_FOUND",
			"Block reference is stale (parent does not contain it)",
		);
	}

	if (opts.dryRun) return respondDryRun(opts);

	const tracked = document.isTrackChangesEnabled();

	if (canPreserveFormatting(spec, opts)) {
		coreApplyFormattingPreservingEdit(
			document,
			blockRef.node,
			spec.text,
			spec.paragraphOptions,
			opts.authorFlag,
			tracked,
		);
		await document.save(opts.outputPath);
		return emitEditAck(opts);
	}

	if (spec.kind === "code") {
		ensureCodeBlockStyles(document, spec.language);
	}
	const newParagraphs = buildNewParagraphs(spec);

	if (tracked) {
		applyTrackedRangeReplace(
			document,
			blockRef.parent,
			targetIndex,
			targetIndex,
			newParagraphs,
			opts.authorFlag,
		);
	} else {
		applyUntrackedRangeReplace(
			blockRef.parent,
			targetIndex,
			targetIndex,
			newParagraphs,
		);
	}

	await document.save(opts.outputPath);
	return emitEditAck(opts);
}

/** The paragraph-content specs that produce one or more new paragraphs. */
type ParagraphContentSpec = Extract<
	EditSpec,
	{ kind: "text" } | { kind: "runs" } | { kind: "code" }
>;

/** The formatting-preservation path applies only to `--text` (not `--runs`,
 *  which already lets the agent specify per-run formatting). It also bows
 *  out when the agent passed any explicit run-level format flag — those
 *  apply uniformly to the new paragraph, which conflicts with per-token
 *  inheritance. `--no-formatting` is the explicit opt-out. */
function canPreserveFormatting(
	spec: ParagraphContentSpec,
	opts: ValidatedOptions,
): spec is Extract<EditSpec, { kind: "text" }> {
	if (opts.noFormatting) return false;
	if (spec.kind !== "text") return false;
	const format = spec.format;
	if (format.color || format.bold || format.italic) return false;
	return true;
}

/** Build the new paragraph(s) for a paragraph-content spec. Text/runs produce
 *  a single paragraph; code produces one paragraph per source line via
 *  `buildCodeBlockParagraphs`. The single-anchor edit path routes a multi-
 *  paragraph result through `applyTrackedRangeReplace` / `applyUntrackedRangeReplace`
 *  with `startIndex === endIndex` (M=1, N=K), so multi-line code lands cleanly. */
function buildNewParagraphs(spec: ParagraphContentSpec): XmlNode[] {
	if (spec.kind === "code") {
		return buildCodeBlockParagraphs(spec.content, spec.language);
	}
	if (spec.kind === "text") {
		return [
			<Paragraph
				text={spec.text}
				{...spec.paragraphOptions}
				{...(spec.format.color ? { color: spec.format.color } : {})}
				{...(spec.format.bold ? { bold: true as const } : {})}
				{...(spec.format.italic ? { italic: true as const } : {})}
			/>,
		];
	}
	return [<Paragraph runs={spec.runs} {...spec.paragraphOptions} />];
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
