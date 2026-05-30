import type { Document } from "../ast/document";
import type { BlockReference } from "../ast/document/body";
import type { Run, SectionType } from "../ast/types";
import { Paragraph, type ParagraphOptions } from "../blocks";
import { buildCodeBlockParagraphs, ensureCodeBlockStyles } from "../code-block";
import { EquationParseError, latexToOmml } from "../equation";
import {
	computeExtentEmu,
	Image,
	type ImageSource,
	ImageSourceError,
	Images,
	loadImageSource,
	nextDrawingId,
} from "../image";
import { w } from "../jsx";
import type { XmlNode } from "../parser";
import { SentinelSectionParagraph } from "../sections";
import { BlankTable, type TableBorders, type TableLayout } from "../table";
import { TrackChanges } from "../track-changes";

/** Cross-cutting lens over "insert a block (of any kind) at a target location."
 * Stateless — `new Insert(document).paragraph(blockRef, spec, …)` builds the
 * block(s), provisions any styles / list ids / image relationships it needs,
 * applies tracked-change wrapping when `<w:trackChanges/>` is on, and returns
 * the freshly-built XmlNodes. **The caller splices and saves** so `--dry-run`
 * can skip persistence cleanly. Throws `InsertError(code, message, hint?)` for
 * domain failures (LaTeX parse, image source, unsupported tracking combo). */
export class Insert {
	constructor(private document: Document) {}

	async paragraph(
		blockRef: BlockReference,
		spec: InsertSpec,
		paragraphOptions: ParagraphOptions,
		opts: { placement: "before" | "after"; authorFlag?: string },
	): Promise<XmlNode[]> {
		// `--task` / `--list` paragraph: resolve the list numId (inherit from the
		// anchor if it's already a list, else allocate fresh of the requested
		// kind). Done first because allocateNum needs the package, and the
		// anchor inherit needs the resolved blockRef.
		const options2 = paragraphOptions as ParagraphOptions & {
			listKind?: "bullet" | "ordered";
			explicitLevel?: number;
		};
		if (options2.taskState !== undefined || options2.list !== undefined) {
			resolveListContext(this.document, blockRef, spec, options2);
			this.document.ensureStyles().ensureStyle("ListParagraph");
		}

		const built = await buildInsertedParagraph(
			this.document,
			spec,
			paragraphOptions,
		);
		this.document.ensureStyles().ensureReferencedStyle(paragraphOptions.style);
		if (spec.kind === "runs") {
			this.document.ensureStyles().ensureReferencedRunStyles(spec.runs);
		}
		if (spec.kind === "code") {
			// Provisions `Code` (character) + `CodeBlock` (paragraph), and when a
			// language was given, the `CodeBlock-LANG` derived paragraph style so
			// the language survives round-trip.
			ensureCodeBlockStyles(this.document, spec.language);
		}
		const blocks = Array.isArray(built) ? built : [built];

		if (this.document.isTrackChangesEnabled()) {
			// Tables under tracking would require per-row <w:trPr><w:ins/> wrappers
			// (ECMA-376 §17.13.5) — deferred. Reject cleanly so the agent knows to
			// toggle tracking off, insert, then back on.
			for (const block of blocks) {
				if (block.tag === "w:tbl") {
					throw new InsertError(
						"TRACKED_CHANGE_CONFLICT",
						"Inserting a table while track-changes is on is not supported",
						"Run `docx track-changes FILE off`, insert the table, then `track-changes on`.",
					);
				}
			}
			const trackChanges = new TrackChanges(this.document);
			for (const block of blocks) {
				trackChanges.applyInsertion(block, opts.authorFlag);
			}
		}

		return blocks;
	}
}

/** Discriminated union covering every insert variant `docx insert` parses out
 * of the CLI flags. The `Insert` lens dispatches on `kind`. */
export type InsertSpec =
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
	  }
	| { kind: "code"; content: string; language?: string }
	| { kind: "equation"; latex: string; display: boolean };

export type TextFormatting = {
	color?: string;
	bold?: boolean;
	italic?: boolean;
};

/** Domain error from `Insert.paragraph`. `code` is a literal subset of the
 * CLI's `ErrorCode` union so callers can `return fail(err.code, err.message,
 * err.hint)` directly — no cast, full type-check coverage. */
export type InsertErrorCode =
	| "USAGE"
	| "TRACKED_CHANGE_CONFLICT"
	| "IMAGE_SOURCE";

export class InsertError extends Error {
	constructor(
		public code: InsertErrorCode,
		message: string,
		public hint?: string,
	) {
		super(message);
		this.name = "InsertError";
	}
}

/** Build the block(s) to insert. Returns either a single XmlNode (one
 * paragraph / one table) or an array (multi-line code blocks produce one
 * `<w:p>` per source line). */
async function buildInsertedParagraph(
	document: Document,
	spec: InsertSpec,
	paragraphOptions: ParagraphOptions,
): Promise<XmlNode | XmlNode[]> {
	switch (spec.kind) {
		case "text":
			return buildTextParagraph(document, spec, paragraphOptions);
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
			return buildImageParagraph(document, spec, paragraphOptions);
		case "code":
			return buildCodeBlockParagraphs(spec.content, spec.language);
		case "equation":
			return buildEquationParagraph(spec, paragraphOptions);
	}
}

/** Build a paragraph whose only content is an equation. Display equations
 * carry the `<m:oMathPara>` wrapper as a direct child of `<w:p>`; inline
 * equations sit alongside text runs (here, alone). */
function buildEquationParagraph(
	spec: Extract<InsertSpec, { kind: "equation" }>,
	paragraphOptions: ParagraphOptions,
): XmlNode {
	let omml: XmlNode;
	try {
		omml = latexToOmml(spec.latex, spec.display);
	} catch (error) {
		if (error instanceof EquationParseError) {
			throw new InsertError(
				"USAGE",
				`Could not parse LaTeX equation: ${error.message}`,
				"Check the LaTeX syntax. The equation goes through temml — it accepts most KaTeX/MathJax LaTeX, but unknown commands fail.",
			);
		}
		throw error;
	}
	return <EquationParagraph omml={omml} {...paragraphOptions} />;
}

/** Wrap an OMML element (`<m:oMath>` for inline, `<m:oMathPara>` for display)
 * in a `<w:p>`. Small JSX component rather than `<Paragraph>` so the
 * `<m:oMath>` child isn't accidentally promoted into a run. */
function EquationParagraph({
	omml,
	style,
	alignment,
}: { omml: XmlNode } & ParagraphOptions): XmlNode {
	return (
		<w.p>
			{(style || alignment) && (
				<w.pPr>
					{style && <w.pStyle w-val={style} />}
					{alignment && <w.jc w-val={alignment} />}
				</w.pPr>
			)}
			{omml}
		</w.p>
	);
}

async function buildImageParagraph(
	document: Document,
	spec: Extract<InsertSpec, { kind: "image" }>,
	paragraphOptions: ParagraphOptions,
): Promise<XmlNode> {
	let source: ImageSource;
	try {
		source = await loadImageSource(spec.src);
	} catch (error) {
		if (error instanceof ImageSourceError) {
			throw new InsertError("IMAGE_SOURCE", error.message);
		}
		throw error;
	}

	const extent = computeExtentEmu(source, {
		widthInches: spec.widthInches,
		heightInches: spec.heightInches,
	});
	if (!extent) {
		throw new InsertError(
			"USAGE",
			`Could not read pixel dimensions from ${spec.src}`,
			"Pass --width INCHES (and optionally --height INCHES) to size it explicitly.",
		);
	}

	const { relationshipId } = new Images(document).add(source);
	const imageRun = (
		<Image
			relationshipId={relationshipId}
			drawingId={nextDrawingId(document.documentTree)}
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
	document: Document,
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
		wrapFirstRunInHyperlink(document, paragraphNode, spec.hyperlinkUrl);
	}
	return paragraphNode;
}

function wrapFirstRunInHyperlink(
	document: Document,
	paragraph: XmlNode,
	url: string,
): void {
	const relationshipId = document.relationships.addHyperlink(url);

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

/** Resolve the numId / level for a `--list` / `--task` insertion. Inherits
 * from the anchor paragraph when it's already in a list; otherwise allocates
 * a fresh numbering of the requested kind. Throws `InsertError` if the spec
 * isn't compatible with list-style content. */
function resolveListContext(
	document: Document,
	blockRef: BlockReference,
	spec: InsertSpec,
	paragraphOptions: ParagraphOptions & {
		listKind?: "bullet" | "ordered";
		explicitLevel?: number;
	},
): void {
	if (spec.kind !== "text" && spec.kind !== "runs") {
		throw new InsertError(
			"USAGE",
			"--task / --list requires --text or --runs (not --code, --image, --table, --section, or break flags)",
		);
	}
	const explicitLevel = paragraphOptions.explicitLevel;
	const kind: "bullet" | "ordered" = paragraphOptions.listKind ?? "bullet";
	// If --list was specified, the sentinel `numId: -1` means "resolve me";
	// otherwise (--task only) the list field is absent at this point.
	const needsResolve =
		!paragraphOptions.list || paragraphOptions.list.numId === -1;
	if (!needsResolve) {
		if (explicitLevel !== undefined && paragraphOptions.list) {
			paragraphOptions.list.level = explicitLevel;
		}
		return;
	}
	const inherited = readListContext(blockRef.node);
	if (inherited) {
		paragraphOptions.list = {
			level: explicitLevel ?? inherited.level,
			numId: inherited.numId,
		};
		return;
	}
	paragraphOptions.list = {
		level: explicitLevel ?? 0,
		numId: document.ensureNumbering().allocate(kind),
	};
}

function readListContext(
	anchor: XmlNode,
): { level: number; numId: number } | null {
	if (anchor.tag !== "w:p") return null;
	const numPr = anchor.findChild("w:pPr")?.findChild("w:numPr");
	if (!numPr) return null;
	// `numId="0"` is the OOXML sentinel for "remove this paragraph from any
	// numbered list" (ECMA-376 §17.9.18) — it's NOT a valid list to inherit.
	// `Number("")` is `0`, so guard explicitly against missing val too.
	const numIdRaw = numPr.findChild("w:numId")?.getAttribute("w:val");
	if (!numIdRaw) return null;
	const numId = Number(numIdRaw);
	if (!Number.isFinite(numId) || numId <= 0) return null;
	const level = Number(numPr.findChild("w:ilvl")?.getAttribute("w:val") ?? "0");
	return { level, numId };
}
