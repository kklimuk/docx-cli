import {
	applyColumns,
	applySectionType,
	type BlockRangeReference,
	type BlockReference,
	type DocView,
	isSectionType,
	isTrackChangesEnabled,
	LocatorParseError,
	LocatorResolveError,
	mintRevisionMeta,
	parseLocator,
	type Run,
	resolveBlockRange,
	type SectionType,
	saveDocView,
	wrapSectPrChange,
} from "@core";
import { Paragraph, type ParagraphOptions } from "@core/blocks";
import {
	buildCodeBlockParagraphs,
	ensureCodeBlockStyles,
} from "@core/code-block";
import type { XmlNode } from "@core/parser";
import { ensureReferencedRunStyles, ensureReferencedStyle } from "@core/styles";
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

	const view = await openOrFail(opts.filePath);
	if (typeof view === "number") return view;

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
		return commitRangeReplacement(view, opts);
	}

	const blockRef = await resolveBlockOrFail(view, opts.locator);
	if (typeof blockRef === "number") return blockRef;

	if (opts.spec.kind === "section") {
		return commitSectionPropertyEdit(view, blockRef, opts.spec, opts);
	}
	ensureReferencedStyle(view, opts.spec.paragraphOptions.style);
	if (opts.spec.kind === "runs") {
		ensureReferencedRunStyles(view, opts.spec.runs);
	}
	return commitParagraphReplacement(view, blockRef, opts.spec, opts);
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
	view: DocView,
	opts: ValidatedOptions,
): Promise<number> {
	if (opts.spec.kind === "section") {
		// Type narrowing — caller already rejected this above.
		return fail("USAGE", "Section edits don't support range locators", HELP);
	}

	let rangeRef: BlockRangeReference;
	try {
		const locator = parseLocator(opts.locator);
		if (locator.kind !== "blockRange") {
			return fail("INVALID_LOCATOR", `Expected pN-pM, got ${opts.locator}`);
		}
		rangeRef = resolveBlockRange(
			view,
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

	ensureReferencedStyle(view, opts.spec.paragraphOptions.style);
	if (opts.spec.kind === "runs") {
		ensureReferencedRunStyles(view, opts.spec.runs);
	}
	if (opts.spec.kind === "code") {
		ensureCodeBlockStyles(view, opts.spec.language);
	}

	const tracked = isTrackChangesEnabled(view);
	if (tracked) {
		const guard = await rejectNonParagraphTrackedRange(rangeRef, opts.locator);
		if (guard !== null) return guard;
	}

	if (opts.dryRun) return respondDryRun(opts);

	const newParagraphs = buildNewParagraphs(opts.spec);
	if (tracked) {
		applyTrackedRangeReplace(
			view,
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

	await saveDocView(view, opts.outputPath);
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

	const contentFlags = [
		text !== undefined,
		runsJson !== undefined,
		codeInline !== undefined,
		codeFile !== undefined,
	].filter(Boolean).length;
	if (contentFlags === 0) {
		return fail(
			"USAGE",
			"Missing content: pass --text, --runs, --code, or --code-file",
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
	view: DocView,
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

	if (isTrackChangesEnabled(view)) {
		wrapSectPrChange(blockRef.node, mintRevisionMeta(view, opts.authorFlag));
	}
	applyColumns(blockRef.node, spec.columns);
	applySectionType(blockRef.node, spec.sectionType);

	await saveDocView(view, opts.outputPath);
	return emitEditAck(opts);
}

async function commitParagraphReplacement(
	view: DocView,
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

	const tracked = isTrackChangesEnabled(view);

	if (canPreserveFormatting(spec, opts)) {
		coreApplyFormattingPreservingEdit(
			view,
			blockRef.node,
			spec.text,
			spec.paragraphOptions,
			opts.authorFlag,
			tracked,
		);
		await saveDocView(view, opts.outputPath);
		return emitEditAck(opts);
	}

	if (spec.kind === "code") {
		ensureCodeBlockStyles(view, spec.language);
	}
	const newParagraphs = buildNewParagraphs(spec);

	if (tracked) {
		applyTrackedRangeReplace(
			view,
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

	await saveDocView(view, opts.outputPath);
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
			(
				<Paragraph
					text={spec.text}
					{...spec.paragraphOptions}
					{...(spec.format.color ? { color: spec.format.color } : {})}
					{...(spec.format.bold ? { bold: true as const } : {})}
					{...(spec.format.italic ? { italic: true as const } : {})}
				/>
			) as XmlNode,
		];
	}
	return [
		(<Paragraph runs={spec.runs} {...spec.paragraphOptions} />) as XmlNode,
	];
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
