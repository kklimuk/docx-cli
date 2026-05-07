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
import { w } from "@core/jsx";
import { XmlNode } from "@core/parser";
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
import { Paragraph, type ParagraphOptions } from "./emit";

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
`;

export async function run(args: string[]): Promise<number> {
	const opts = await parseAndValidateOptions(args);
	if (typeof opts === "number") return opts;

	const view = await openOrFail(opts.filePath);
	if (typeof view === "number") return view;

	const blockRef = await resolveBlockOrFail(view, opts.placement.locator);
	if (typeof blockRef === "number") return blockRef;

	const paragraph = buildInsertedParagraph(
		view,
		opts.spec,
		opts.paragraphOptions,
	);
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
	| { kind: "section"; columns?: number; sectionType?: SectionType };

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

async function chooseContentSpec(
	values: RawValues,
): Promise<InsertSpec | number> {
	const text = values.text as string | undefined;
	const runsJson = values.runs as string | undefined;
	const url = values.url as string | undefined;
	const pageBreak = (values["page-break"] as boolean | undefined) ?? false;
	const columnBreak = (values["column-break"] as boolean | undefined) ?? false;
	const sectionFlag = (values.section as boolean | undefined) ?? false;
	const columnsRaw = values.columns as string | undefined;
	const sectionTypeRaw = values.type as string | undefined;

	const contentFlagCount =
		(text !== undefined ? 1 : 0) +
		(runsJson !== undefined ? 1 : 0) +
		(pageBreak ? 1 : 0) +
		(columnBreak ? 1 : 0) +
		(sectionFlag ? 1 : 0);
	if (contentFlagCount === 0) {
		return fail(
			"USAGE",
			"Missing content: pass --text, --runs, --page-break, --column-break, or --section",
			HELP,
		);
	}
	if (contentFlagCount > 1) {
		return fail(
			"USAGE",
			"Pass only one of --text, --runs, --page-break, --column-break, --section",
			HELP,
		);
	}
	if (url !== undefined && text === undefined) {
		return fail("USAGE", "--url requires --text", HELP);
	}
	if (
		(columnsRaw !== undefined || sectionTypeRaw !== undefined) &&
		!sectionFlag
	) {
		return fail("USAGE", "--columns and --type require --section", HELP);
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
			...(url ? { hyperlinkUrl: url } : {}),
		};
	}

	if (runsJson !== undefined) {
		const runs = await parseRunsArg(runsJson);
		if (typeof runs === "number") return runs;
		return { kind: "runs", runs };
	}

	if (pageBreak) return { kind: "break", breakKind: "page" };
	if (columnBreak) return { kind: "break", breakKind: "column" };

	const sectionFlags = await parseSectionFlags(values);
	if (typeof sectionFlags === "number") return sectionFlags;
	return { kind: "section", ...sectionFlags };
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

function buildInsertedParagraph(
	view: DocView,
	spec: InsertSpec,
	paragraphOptions: ParagraphOptions,
): XmlNode {
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
	}
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
			{...(spec.format.color ? { color: spec.format.color } : {})}
			{...(spec.format.bold ? { bold: true as const } : {})}
			{...(spec.format.italic ? { italic: true as const } : {})}
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
