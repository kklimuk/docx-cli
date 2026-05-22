import {
	applyColumns,
	applySectionType,
	type BlockReference,
	convertTextToDelText,
	createRevisionAllocator,
	Del,
	type DocView,
	Ins,
	isSectionType,
	isTrackChangesEnabled,
	type Run,
	resolveAuthor,
	resolveDate,
	type SectionType,
	saveDocView,
	type TrackedMeta,
	wrapSectPrChange,
} from "@core";
import { Paragraph, type ParagraphOptions } from "@core/blocks";
import { isRunBearingWrapper, XmlNode } from "@core/parser";
import { ensureReferencedStyle } from "@core/styles";
import { parseArgs } from "util";
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
import {
	buildTrackedRuns,
	buildUntrackedRuns,
	diffTokens,
	extractOldTokens,
	tokenize,
} from "./preserve-formatting";

const HELP = `docx edit — replace a paragraph or modify a section at a locator

Usage:
  docx edit FILE [options]

Locator (required):
  --at LOCATOR      Block to edit (paragraph pN or section sN)

Paragraph content (one required for paragraph locators):
  --text TEXT       Replace with a single-run paragraph
  --runs JSON       Replace with custom runs (Run[] JSON)

Paragraph options:
  --style NAME       Paragraph style (e.g., Heading1)
  --alignment ALIGN  left | center | right | justify

Run options (only with --text):
  --color HEX       Run color, hex (e.g., 800080 for purple)
  --bold            Bold
  --italic          Italic

Section options (for section locators sN):
  --columns N        Number of columns for the targeted section
  --type T           continuous | nextPage | evenPage | oddPage | nextColumn

Formatting (for --text):
  By default --text preserves run-level formatting (bold/italic/color/etc.)
  on words shared between the old and new text via a word-level diff. New
  words inherit formatting from the nearest unchanged neighbor. Pass
  --no-formatting to fall back to a single fresh run with no formatting.
  Passing --color/--bold/--italic also bypasses preservation — those flags
  apply uniformly to the new paragraph.

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
  docx edit doc.docx --at s0 --columns 2 --type continuous
`;

export async function run(args: string[]): Promise<number> {
	const opts = await parseAndValidateOptions(args);
	if (typeof opts === "number") return opts;

	const view = await openOrFail(opts.filePath);
	if (typeof view === "number") return view;

	const blockRef = await resolveBlockOrFail(view, opts.locator);
	if (typeof blockRef === "number") return blockRef;

	if (opts.spec.kind === "section") {
		return commitSectionPropertyEdit(view, blockRef, opts.spec, opts);
	}
	ensureReferencedStyle(view, opts.spec.paragraphOptions.style);
	return commitParagraphReplacement(view, blockRef, opts.spec, opts);
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
	| { kind: "runs"; runs: Run[]; paragraphOptions: ParagraphOptions };

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
	if (!text && !runsJson) {
		return fail("USAGE", "Missing content: pass --text or --runs", HELP);
	}
	if (text && runsJson) {
		return fail("USAGE", "Pass either --text or --runs, not both", HELP);
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

	const runs = await parseRunsArg(runsJson as string);
	if (typeof runs === "number") return runs;
	return { kind: "runs", runs, paragraphOptions };
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
		const allocator = createRevisionAllocator(view);
		const meta: TrackedMeta = {
			author: resolveAuthor(opts.authorFlag),
			date: resolveDate(),
			revisionId: allocator.next(),
		};
		wrapSectPrChange(blockRef.node, meta);
	}
	applyColumns(blockRef.node, spec.columns);
	applySectionType(blockRef.node, spec.sectionType);

	await saveDocView(view, opts.outputPath);
	return emitEditAck(opts);
}

async function commitParagraphReplacement(
	view: DocView,
	blockRef: BlockReference,
	spec: Extract<EditSpec, { kind: "text" } | { kind: "runs" }>,
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
		applyFormattingPreservingEdit(
			view,
			blockRef.node,
			spec,
			opts.authorFlag,
			tracked,
		);
		await saveDocView(view, opts.outputPath);
		return emitEditAck(opts);
	}

	const newParagraph = buildReplacementParagraph(spec);

	if (tracked) {
		applyTrackedEdit(view, blockRef.node, newParagraph, opts.authorFlag);
	} else {
		blockRef.parent.splice(targetIndex, 1, newParagraph);
	}

	await saveDocView(view, opts.outputPath);
	return emitEditAck(opts);
}

/** The formatting-preservation path applies only to `--text` (not `--runs`,
 *  which already lets the agent specify per-run formatting). It also bows
 *  out when the agent passed any explicit run-level format flag — those
 *  apply uniformly to the new paragraph, which conflicts with per-token
 *  inheritance. `--no-formatting` is the explicit opt-out. */
function canPreserveFormatting(
	spec: Extract<EditSpec, { kind: "text" } | { kind: "runs" }>,
	opts: ValidatedOptions,
): spec is Extract<EditSpec, { kind: "text" }> {
	if (opts.noFormatting) return false;
	if (spec.kind !== "text") return false;
	const format = spec.format;
	if (format.color || format.bold || format.italic) return false;
	return true;
}

function applyFormattingPreservingEdit(
	view: DocView,
	existingParagraph: XmlNode,
	spec: Extract<EditSpec, { kind: "text" } | { kind: "runs" }> & {
		kind: "text";
	},
	authorFlag: string | undefined,
	tracked: boolean,
): void {
	const oldTokens = extractOldTokens(existingParagraph);
	const newTokens = tokenize(spec.text);
	const ops = diffTokens(oldTokens, newTokens);

	let runChildren: XmlNode[];
	if (tracked) {
		const allocator = createRevisionAllocator(view);
		const baseMeta = { author: resolveAuthor(authorFlag), date: resolveDate() };
		const mintMeta = (): TrackedMeta => ({
			...baseMeta,
			revisionId: allocator.next(),
		});
		runChildren = buildTrackedRuns(ops, mintMeta);
	} else {
		runChildren = buildUntrackedRuns(ops);
	}

	// Preserve <w:pPr>, bookmarks, comment markers, and other non-run-
	// bearing children. Strip top-level <w:r> AND every run-bearing
	// wrapper (`<w:ins>`, `<w:del>`, `<w:moveFrom>`, `<w:moveTo>`,
	// `<w:hyperlink>`, `<w:fldSimple>`, `<w:smartTag>`) — `extractOldTokens`
	// already flattened the visible-in-accepted-view wrappers into the
	// diff input, and the new diff runs replace them. Without this strip
	// the wrappers' inner text would leak into the output alongside the
	// diff runs (the bug from agent feedback: chained edits + hyperlinked
	// paragraphs duplicating text).
	// Optional --style / --alignment tweaks the existing pPr in place.
	const rebuilt: XmlNode[] = [];
	for (const child of existingParagraph.children) {
		if (child.tag === "w:r") continue;
		if (isRunBearingWrapper(child.tag)) continue;
		rebuilt.push(child);
	}
	applyParagraphOptionsInPlace(rebuilt, spec.paragraphOptions);
	rebuilt.push(...runChildren);
	existingParagraph.children = rebuilt;
}

/** Apply `--style` / `--alignment` to the paragraph's existing `<w:pPr>`,
 *  creating one if needed. Mirrors the simple-emit path's behavior without
 *  building a fresh paragraph. */
function applyParagraphOptionsInPlace(
	rebuilt: XmlNode[],
	options: ParagraphOptions,
): void {
	if (!options.style && !options.alignment) return;
	let pPr = rebuilt.find((child) => child.tag === "w:pPr");
	if (!pPr) {
		pPr = new XmlNode("w:pPr");
		rebuilt.unshift(pPr);
	}
	if (options.style) {
		const existingStyle = pPr.findChild("w:pStyle");
		if (existingStyle) {
			existingStyle.setAttribute("w:val", options.style);
		} else {
			const styleNode = new XmlNode("w:pStyle", { "w:val": options.style });
			pPr.children.unshift(styleNode);
		}
	}
	if (options.alignment) {
		const existingJc = pPr.findChild("w:jc");
		if (existingJc) {
			existingJc.setAttribute("w:val", options.alignment);
		} else {
			pPr.children.push(new XmlNode("w:jc", { "w:val": options.alignment }));
		}
	}
}

function buildReplacementParagraph(
	spec: Extract<EditSpec, { kind: "text" } | { kind: "runs" }>,
): XmlNode {
	if (spec.kind === "text") {
		return (
			<Paragraph
				text={spec.text}
				{...spec.paragraphOptions}
				{...(spec.format.color ? { color: spec.format.color } : {})}
				{...(spec.format.bold ? { bold: true as const } : {})}
				{...(spec.format.italic ? { italic: true as const } : {})}
			/>
		);
	}
	return <Paragraph runs={spec.runs} {...spec.paragraphOptions} />;
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

function applyTrackedEdit(
	view: DocView,
	existingParagraph: XmlNode,
	newParagraph: XmlNode,
	authorFlag: string | undefined,
): void {
	const allocator = createRevisionAllocator(view);
	const baseMeta = { author: resolveAuthor(authorFlag), date: resolveDate() };
	const mintMeta = (): TrackedMeta => ({
		...baseMeta,
		revisionId: allocator.next(),
	});

	// Treat run-bearing wrappers as part of the deletable old content:
	// flatten their `<w:r>` children into oldRuns. Without this, prior
	// `<w:ins>`/`<w:del>`/`<w:hyperlink>`/etc. wrappers would persist in
	// `oldNonRuns` and leak alongside the new del+ins blocks (the bug
	// from agent feedback: chained edits + hyperlinked paragraphs
	// duplicating text).
	const oldRuns: XmlNode[] = [];
	const oldNonRuns: XmlNode[] = [];
	for (const child of existingParagraph.children) {
		if (child.tag === "w:r") {
			oldRuns.push(child);
			continue;
		}
		if (isRunBearingWrapper(child.tag)) {
			collectRunsFromWrapper(child, oldRuns);
			continue;
		}
		oldNonRuns.push(child);
	}

	let newPPr: XmlNode | null = null;
	const newRuns: XmlNode[] = [];
	for (const child of newParagraph.children) {
		if (child.tag === "w:pPr") newPPr = child;
		else if (child.tag === "w:r") newRuns.push(child);
	}

	const rebuilt: XmlNode[] = [];
	if (newPPr) {
		rebuilt.push(newPPr);
		for (const child of oldNonRuns) {
			if (child.tag !== "w:pPr") rebuilt.push(child);
		}
	} else {
		rebuilt.push(...oldNonRuns);
	}
	if (oldRuns.length > 0) {
		const deletedRuns = oldRuns.map((run) => convertTextToDelText(run));
		rebuilt.push(<Del meta={mintMeta()}>{deletedRuns}</Del>);
	}
	if (newRuns.length > 0) {
		rebuilt.push(<Ins meta={mintMeta()}>{newRuns}</Ins>);
	}
	existingParagraph.children = rebuilt;
}

/** Recursively collect every `<w:r>` from inside a run-bearing wrapper.
 *  Used by the legacy whole-paragraph tracked edit path so that prior
 *  tracked-change wrappers don't leak into the rebuilt paragraph
 *  alongside the new del+ins blocks. */
function collectRunsFromWrapper(wrapper: XmlNode, out: XmlNode[]): void {
	for (const child of wrapper.children) {
		if (child.tag === "w:r") {
			out.push(child);
			continue;
		}
		if (isRunBearingWrapper(child.tag)) {
			collectRunsFromWrapper(child, out);
		}
	}
}
