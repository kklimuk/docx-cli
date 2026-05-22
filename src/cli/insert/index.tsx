import {
	addHyperlinkRelationship,
	type BlockReference,
	createRevisionAllocator,
	type DocView,
	Ins,
	isSectionType,
	isTrackChangesEnabled,
	markParagraphMarkAs,
	type Run,
	resolveAuthor,
	resolveDate,
	type SectionType,
	SentinelSectionParagraph,
	saveDocView,
	type TrackedMeta,
} from "@core";
import { Paragraph, type ParagraphOptions } from "@core/blocks";
import {
	addImagePart,
	computeExtentEmu,
	Image,
	type ImageSource,
	ImageSourceError,
	loadImageSource,
	nextDrawingId,
} from "@core/image";
import { w } from "@core/jsx";
import { XmlNode } from "@core/parser";
import { ensureReferencedStyle } from "@core/styles";
import { BlankTable, type TableBorders, type TableLayout } from "@core/table";
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

Paragraph options:
  --style NAME       Apply paragraph style (e.g., Heading1)
  --alignment ALIGN  left | center | right | justify

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
`;

export async function run(args: string[]): Promise<number> {
	const opts = await parseAndValidateOptions(args);
	if (typeof opts === "number") return opts;

	const view = await openOrFail(opts.filePath);
	if (typeof view === "number") return view;

	const blockRef = await resolveBlockOrFail(view, opts.placement.locator);
	if (typeof blockRef === "number") return blockRef;

	const paragraph = await buildInsertedParagraph(
		view,
		opts.spec,
		opts.paragraphOptions,
	);
	if (typeof paragraph === "number") return paragraph;
	ensureReferencedStyle(view, opts.paragraphOptions.style);
	return commitInsertedParagraph(view, blockRef, paragraph, opts);
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

	const placement = await parseTargetPlacement(parsed.values);
	if (typeof placement === "number") return placement;

	const spec = await chooseContentSpec(parsed.values);
	if (typeof spec === "number") return spec;

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

type InsertSpec =
	| {
			kind: "text";
			text: string;
			format: TextFormatting;
			hyperlinkUrl?: string;
	  }
	| { kind: "runs"; runs: Run[] }
	| { kind: "break"; breakKind: "page" | "column" }
	| { kind: "section"; columns?: number; sectionType?: SectionType }
	| {
			kind: "table";
			rows: number;
			cols: number;
			widths?: number[];
			tableWidth?: { value: number; unit: "dxa" | "pct" };
			borders?: TableBorders;
			layout?: TableLayout;
	  }
	| {
			kind: "image";
			src: string;
			alt?: string;
			widthInches?: number;
			heightInches?: number;
	  };

type TextFormatting = {
	color?: string;
	bold?: boolean;
	italic?: boolean;
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
	for (const kind of CONTENT_KINDS) {
		if (kind.flag === chosen.flag) continue;
		const orphan = kind.subFlags.find((flag) => values[flag] !== undefined);
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

async function buildInsertedParagraph(
	view: DocView,
	spec: InsertSpec,
	paragraphOptions: ParagraphOptions,
): Promise<XmlNode | number> {
	switch (spec.kind) {
		case "text":
			return buildTextParagraph(view, spec, paragraphOptions);
		case "runs":
			return <Paragraph runs={spec.runs} {...paragraphOptions} />;
		case "break":
			return (
				<Paragraph
					runs={[{ type: "break", kind: spec.breakKind }]}
					{...paragraphOptions}
				/>
			);
		case "section":
			return (
				<SentinelSectionParagraph
					{...(spec.columns !== undefined ? { columns: spec.columns } : {})}
					{...(spec.sectionType ? { sectionType: spec.sectionType } : {})}
				/>
			);
		case "table":
			return (
				<BlankTable
					rows={spec.rows}
					cols={spec.cols}
					widths={spec.widths}
					width={spec.tableWidth}
					borders={spec.borders}
					layout={spec.layout}
				/>
			);
		case "image":
			return buildImageParagraph(view, spec, paragraphOptions);
	}
}

async function buildImageParagraph(
	view: DocView,
	spec: Extract<InsertSpec, { kind: "image" }>,
	paragraphOptions: ParagraphOptions,
): Promise<XmlNode | number> {
	let source: ImageSource;
	try {
		source = await loadImageSource(spec.src);
	} catch (error) {
		if (error instanceof ImageSourceError) {
			return fail("IMAGE_SOURCE", error.message);
		}
		throw error;
	}

	const extent = computeExtentEmu(source, {
		widthInches: spec.widthInches,
		heightInches: spec.heightInches,
	});
	if (!extent) {
		return fail(
			"USAGE",
			`Could not read pixel dimensions from ${spec.src}`,
			"Pass --width INCHES (and optionally --height INCHES) to size it explicitly.",
		);
	}

	const { relationshipId } = addImagePart(view, source);
	const imageRun = (
		<Image
			relationshipId={relationshipId}
			drawingId={nextDrawingId(view.documentTree)}
			widthEmu={extent.widthEmu}
			heightEmu={extent.heightEmu}
			alt={spec.alt}
		/>
	);

	const { style, alignment } = paragraphOptions;
	return (
		<w.p>
			{style || alignment ? (
				<w.pPr>
					{style ? <w.pStyle w-val={style} /> : null}
					{alignment ? <w.jc w-val={alignment} /> : null}
				</w.pPr>
			) : null}
			{imageRun}
		</w.p>
	);
}

function buildTextParagraph(
	view: DocView,
	spec: Extract<InsertSpec, { kind: "text" }>,
	paragraphOptions: ParagraphOptions,
): XmlNode {
	const paragraphNode = (
		<Paragraph
			text={spec.text}
			{...paragraphOptions}
			color={spec.format.color}
			bold={spec.format.bold}
			italic={spec.format.italic}
		/>
	);
	if (spec.hyperlinkUrl) {
		wrapFirstRunInHyperlink(view, paragraphNode, spec.hyperlinkUrl);
	}
	return paragraphNode;
}

async function commitInsertedParagraph(
	view: DocView,
	blockRef: BlockReference,
	paragraph: XmlNode,
	opts: ValidatedOptions,
): Promise<number> {
	const targetIndex = blockRef.parent.indexOf(blockRef.node);
	if (targetIndex === -1) {
		return fail(
			"BLOCK_NOT_FOUND",
			"Block reference is stale (parent does not contain it)",
		);
	}
	const insertIndex =
		opts.placement.mode === "after" ? targetIndex + 1 : targetIndex;

	if (isTrackChangesEnabled(view)) {
		// Tables under tracking would require per-row <w:trPr><w:ins/> wrappers
		// (ECMA-376 §17.13.5) — defer to S4b. Reject cleanly so the agent
		// knows to toggle tracking off, insert, then back on.
		if (paragraph.tag === "w:tbl") {
			return fail(
				"TRACKED_CHANGE_CONFLICT",
				"Inserting a table while track-changes is on is not supported",
				"Run `docx track-changes FILE off`, insert the table, then `track-changes on`.",
			);
		}
		applyTrackedInsertion(paragraph, view, opts.authorFlag);
	}

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

	blockRef.parent.splice(insertIndex, 0, paragraph);
	await saveDocView(view, opts.outputPath);

	await respondAck({
		ok: true,
		operation: "insert",
		path: opts.outputPath ?? opts.filePath,
		locator: opts.placement.locator,
		placement: opts.placement.mode,
	});
	return EXIT.OK;
}

function wrapFirstRunInHyperlink(
	view: DocView,
	paragraph: XmlNode,
	url: string,
): void {
	const relationships = XmlNode.findRoot(
		view.relationshipsTree,
		"Relationships",
	);
	if (!relationships) {
		throw new Error("Missing <Relationships> root in document rels");
	}
	const relationshipId = addHyperlinkRelationship(relationships, url);
	view.hyperlinksByRelationshipId.set(relationshipId, { url });

	const newChildren: XmlNode[] = [];
	let wrapped = false;
	for (const child of paragraph.children) {
		if (!wrapped && child.tag === "w:r") {
			const wrapper = (
				<w.hyperlink {...{ "r:id": relationshipId }}>{child}</w.hyperlink>
			);
			newChildren.push(wrapper);
			wrapped = true;
			continue;
		}
		newChildren.push(child);
	}
	paragraph.children = newChildren;
}

function applyTrackedInsertion(
	paragraph: XmlNode,
	view: DocView,
	authorFlag: string | undefined,
): void {
	const allocator = createRevisionAllocator(view);
	const baseMeta = { author: resolveAuthor(authorFlag), date: resolveDate() };
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
