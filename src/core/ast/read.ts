import { getListFormat } from "../numbering";
import { XmlNode } from "../parser";
import { readSectionProperties } from "../sections";
import { detectTaskListState } from "../task-list";
import type { DocView } from "./doc-view";
import { decodeSym } from "./sym";
import type {
	Block,
	ChartRun,
	Comment,
	CommentAnchor,
	Doc,
	DocProperties,
	Footnote,
	Hyperlink,
	ImageRun,
	Paragraph,
	Run,
	SectionBreak,
	Table,
	TableCell,
	TableRow,
	TableWidth,
	TextRun,
	TrackedChange,
} from "./types";

const RELATIONSHIP_NAMESPACE_IMAGE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const RELATIONSHIP_NAMESPACE_HYPERLINK =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

type WalkState = {
	imageIndex: number;
	hyperlinkIndex: number;
	trackedChangeIndex: number;
	commentAnchors: Map<string, CommentAnchor>;
	openComments: Map<string, { blockId: string; offset: number }>;
};

export function buildDoc(view: DocView, path: string): Doc {
	readRelationships(view);
	const properties = readDocProperties(view);

	const documentRoot = XmlNode.findRoot(view.documentTree, "w:document");
	if (!documentRoot) {
		throw new Error("Invalid .docx: missing <w:document>");
	}
	const body = documentRoot.findChild("w:body");
	if (!body) {
		throw new Error("Invalid .docx: missing <w:body>");
	}

	const state: WalkState = {
		imageIndex: 0,
		hyperlinkIndex: 0,
		trackedChangeIndex: 0,
		commentAnchors: new Map(),
		openComments: new Map(),
	};

	const blocks = readBlocks(view, body, state);
	const comments = readComments(view, state.commentAnchors);
	const footnotes = readNotes(
		view.footnotesTree,
		"w:footnotes",
		"w:footnote",
		"fn",
	);
	const endnotes = readNotes(
		view.endnotesTree,
		"w:endnotes",
		"w:endnote",
		"en",
	);

	return {
		schemaVersion: 1,
		path,
		properties,
		blocks,
		comments,
		footnotes,
		endnotes,
	};
}

function readBlocks(view: DocView, body: XmlNode, state: WalkState): Block[] {
	const blocks: Block[] = [];
	let paragraphIndex = 0;
	let tableIndex = 0;
	let sectionIndex = 0;

	for (const child of body.children) {
		if (child.tag === "w:p") {
			const id = `p${paragraphIndex++}`;
			blocks.push(readParagraph(view, child, id, state));
			view.blockReferences.set(id, { node: child, parent: body.children });
			const inlineSectPr = findInlineSectPr(child);
			if (inlineSectPr) {
				const sectionId = `s${sectionIndex++}`;
				blocks.push(buildSectionBreak(sectionId, inlineSectPr.node));
				view.blockReferences.set(sectionId, inlineSectPr);
				registerSectPrChange(view, inlineSectPr.node, state, sectionId);
			}
			registerParagraphMarkChanges(view, child, state, id);
			continue;
		}
		if (child.tag === "w:tbl") {
			const id = `t${tableIndex++}`;
			blocks.push(readTable(view, child, id, state));
			view.blockReferences.set(id, { node: child, parent: body.children });
			continue;
		}
		if (child.tag === "w:sectPr") {
			const id = `s${sectionIndex++}`;
			blocks.push(buildSectionBreak(id, child));
			view.blockReferences.set(id, { node: child, parent: body.children });
			registerSectPrChange(view, child, state, id);
		}
	}
	return blocks;
}

/** Register paragraph-mark `<w:ins>` / `<w:del>` markers (lives inside
 * `<w:pPr><w:rPr>`) as tracked-change references, so `track-changes list`
 * surfaces them and `track-changes accept --at tcN` can target them. The
 * blockId is the owning paragraph's pN. Called AFTER `registerSectPrChange`
 * so the tcN ordering matches `cli/track-changes/apply.ts collectTrackedChanges`
 * (run-level → sectPrChange → paragraph-mark, per paragraph). */
function registerParagraphMarkChanges(
	view: DocView,
	paragraph: XmlNode,
	state: WalkState,
	blockId: string,
): void {
	const pPr = paragraph.findChild("w:pPr");
	if (!pPr) return;
	const rPr = pPr.findChild("w:rPr");
	if (!rPr) return;
	for (const child of rPr.children) {
		if (child.tag !== "w:ins" && child.tag !== "w:del") continue;
		const id = `tc${state.trackedChangeIndex++}`;
		view.trackedChangeReferences.set(id, {
			node: child,
			parent: rPr.children,
			blockId,
		});
	}
}

/** Register a <w:sectPrChange> element (if present inside the sectPr) as
 * a tracked-change reference, so `track-changes list/accept/reject` can
 * address it as tcN. The blockId is the corresponding section's sN (both
 * inline and trailing sectPrs). */
function registerSectPrChange(
	view: DocView,
	sectPr: XmlNode,
	state: WalkState,
	blockId: string,
): void {
	const change = sectPr.findChild("w:sectPrChange");
	if (!change) return;
	const id = `tc${state.trackedChangeIndex++}`;
	view.trackedChangeReferences.set(id, {
		node: change,
		parent: sectPr.children,
		blockId,
	});
}

function findInlineSectPr(
	paragraph: XmlNode,
): { node: XmlNode; parent: XmlNode[] } | undefined {
	const pPr = paragraph.findChild("w:pPr");
	if (!pPr) return undefined;
	const sectPr = pPr.findChild("w:sectPr");
	if (!sectPr) return undefined;
	return { node: sectPr, parent: pPr.children };
}

function buildSectionBreak(id: string, sectPr: XmlNode): SectionBreak {
	const props = readSectionProperties(sectPr.children);
	const block: SectionBreak = { id, type: "sectionBreak" };
	if (props.columns !== undefined) block.columns = props.columns;
	if (props.sectionType !== undefined) block.sectionType = props.sectionType;
	return block;
}

function readParagraph(
	view: DocView,
	node: XmlNode,
	id: string,
	state: WalkState,
): Paragraph {
	const paragraph: Paragraph = { id, type: "paragraph", runs: [] };
	const paragraphProperties = node.findChild("w:pPr");
	if (paragraphProperties) {
		applyParagraphProperties(view, paragraph, paragraphProperties);
	}

	const skipNodes = detectTaskListState(
		view,
		paragraph,
		node,
		(sdt, parent) => {
			// Register a checkboxToggle reference at the current walk position so
			// `tcN` ids agree with the apply walker in `cli/track-changes/apply.ts`.
			// The apply walker sees the SDT as the first non-pPr child and treats
			// the toggle as the first tcN of this paragraph; we mirror that here.
			const tcId = `tc${state.trackedChangeIndex++}`;
			view.trackedChangeReferences.set(tcId, {
				node: sdt,
				parent,
				blockId: id,
				kind: "checkboxToggle",
			});
		},
	);

	const context: WalkContext = {
		view,
		blockId: id,
		paragraph,
		activeComments: new Set<string>(),
		state,
		offsetRef: { value: 0 },
		skipNodes,
	};
	walkRunContainer(context, node, undefined, undefined);

	return paragraph;
}

type WalkContext = {
	view: DocView;
	blockId: string;
	paragraph: Paragraph;
	activeComments: Set<string>;
	state: WalkState;
	offsetRef: { value: number };
	skipNodes: Set<XmlNode>;
};

function walkRunContainer(
	context: WalkContext,
	container: XmlNode,
	trackedChange: TrackedChange | undefined,
	hyperlink: Hyperlink | undefined,
): void {
	for (const child of container.children) {
		if (child.tag === "w:pPr") continue;
		if (context.skipNodes.has(child)) continue;
		// Skip ALL `<w:sdt>` content control bodies, not just leading checkbox
		// SDTs (which `detectTaskListState` already pushed into `skipNodes`).
		// The apply walker in `cli/track-changes/apply.ts::visitRunContainer`
		// does the same, and the alignment invariant in the track-changes
		// CLAUDE.md requires both walkers to register `tcN` ids in matching
		// order. Recursing into a Plain Text / Dropdown / Rich Text content
		// control here would surface its nested `<w:ins>`/`<w:del>` as run-
		// level tcNs that apply.ts ignores — drift. The structural-ins/del
		// gap for SDTs is documented in `core/task-list/CLAUDE.md`.
		if (child.tag === "w:sdt") continue;

		if (child.tag === "w:commentRangeStart") {
			const commentId = child.getAttribute("w:id");
			if (commentId) {
				const key = `c${commentId}`;
				context.activeComments.add(key);
				context.state.openComments.set(key, {
					blockId: context.blockId,
					offset: context.offsetRef.value,
				});
			}
			continue;
		}

		if (child.tag === "w:commentRangeEnd") {
			const commentId = child.getAttribute("w:id");
			if (commentId) {
				const key = `c${commentId}`;
				context.activeComments.delete(key);
				const opened = context.state.openComments.get(key);
				if (opened) {
					context.state.commentAnchors.set(key, {
						startBlockId: opened.blockId,
						startOffset: opened.offset,
						endBlockId: context.blockId,
						endOffset: context.offsetRef.value,
					});
					context.state.openComments.delete(key);
				}
			}
			continue;
		}

		if (child.tag === "w:r") {
			const runs = readRun(
				context.view,
				child,
				context.activeComments,
				trackedChange,
				hyperlink,
				context.state,
			);
			for (const run of runs) {
				if (run.type === "text") context.offsetRef.value += run.text.length;
				context.paragraph.runs.push(run);
			}
			continue;
		}

		if (child.tag === "m:oMath") {
			const text = collectMathText(child);
			if (text.length > 0) {
				context.paragraph.runs.push({
					type: "equation",
					text,
					display: false,
				});
			}
			continue;
		}

		if (child.tag === "m:oMathPara") {
			const text = collectMathText(child);
			if (text.length > 0) {
				context.paragraph.runs.push({
					type: "equation",
					text,
					display: true,
				});
			}
			continue;
		}

		if (
			child.tag === "w:ins" ||
			child.tag === "w:del" ||
			child.tag === "w:moveFrom" ||
			child.tag === "w:moveTo"
		) {
			const trackedChangeId = `tc${context.state.trackedChangeIndex++}`;
			const change: TrackedChange = {
				id: trackedChangeId,
				kind: TRACKED_CHANGE_KIND_BY_TAG[child.tag],
				author: child.getAttribute("w:author") ?? "",
				date: child.getAttribute("w:date") ?? "",
				revisionId: child.getAttribute("w:id") ?? "",
			};
			context.view.trackedChangeReferences.set(trackedChangeId, {
				node: child,
				parent: container.children,
				blockId: context.blockId,
			});
			walkRunContainer(context, child, change, hyperlink);
			continue;
		}

		if (child.tag === "w:hyperlink") {
			const link = readHyperlinkProperties(
				context.view,
				child,
				container.children,
				context.state,
			);
			walkRunContainer(context, child, trackedChange, link);
			continue;
		}

		// Transparent wrappers — recurse but contribute no AST node themselves.
		// <w:fldSimple>: a self-contained field; its result text lives in inner
		//   <w:r> / <w:t> children. Dropping it would lose the rendered text.
		// <w:smartTag>: Word's semantic tagging for proper nouns / dates;
		//   contains plain runs.
		if (child.tag === "w:fldSimple" || child.tag === "w:smartTag") {
			walkRunContainer(context, child, trackedChange, hyperlink);
		}
	}
}

const TRACKED_CHANGE_KIND_BY_TAG: Record<
	"w:ins" | "w:del" | "w:moveFrom" | "w:moveTo",
	TrackedChange["kind"]
> = {
	"w:ins": "ins",
	"w:del": "del",
	"w:moveFrom": "moveFrom",
	"w:moveTo": "moveTo",
};

function readHyperlinkProperties(
	view: DocView,
	node: XmlNode,
	parent: XmlNode[],
	state: WalkState,
): Hyperlink | undefined {
	const id = `link${state.hyperlinkIndex++}`;
	const link: Hyperlink = { id };
	const relationshipId = node.getAttribute("r:id");
	if (relationshipId) {
		const relationship = view.hyperlinksByRelationshipId.get(relationshipId);
		if (relationship?.url) link.url = relationship.url;
	}
	const anchor = node.getAttribute("w:anchor");
	if (anchor) link.anchor = anchor;
	const tooltip = node.getAttribute("w:tooltip");
	if (tooltip) link.tooltip = tooltip;
	if (!link.url && !link.anchor && !link.tooltip) {
		state.hyperlinkIndex--;
		return undefined;
	}
	view.hyperlinkById.set(id, {
		node,
		parent,
		...(relationshipId ? { relationshipId } : {}),
	});
	return link;
}

function applyParagraphProperties(
	view: DocView,
	paragraph: Paragraph,
	paragraphProperties: XmlNode,
): void {
	const styleNode = paragraphProperties.findChild("w:pStyle");
	if (styleNode) {
		const value = styleNode.getAttribute("w:val");
		if (value) paragraph.style = value;
	}

	const justification = paragraphProperties.findChild("w:jc");
	if (justification) {
		const value = justification.getAttribute("w:val");
		if (
			value === "left" ||
			value === "center" ||
			value === "right" ||
			value === "justify"
		) {
			paragraph.alignment = value;
		}
	}

	const numberingProperties = paragraphProperties.findChild("w:numPr");
	if (numberingProperties) {
		const indentLevel = numberingProperties.findChild("w:ilvl");
		const numberingId = numberingProperties.findChild("w:numId");
		const level = indentLevel
			? Number(indentLevel.getAttribute("w:val") ?? "0")
			: 0;
		const id = numberingId ? (numberingId.getAttribute("w:val") ?? "") : "";
		const list: { level: number; numId: string; ordered?: boolean } = {
			level,
			numId: id,
		};
		// Resolve the level's numFmt via numbering.xml so the renderer knows
		// whether to emit `1. ` (ordered) or `- ` (bullet) in markdown.
		const format = id ? getListFormat(view, id, level) : undefined;
		if (format !== undefined && format !== "bullet" && format !== "none") {
			list.ordered = true;
		}
		paragraph.list = list;
	}
}

/** A single <w:r> can contain a mix of <w:t>, <w:tab>, <w:br>, <w:drawing>,
 * footnote/endnote refs in any order. We emit one AST Run per child in
 * document order; consecutive <w:t>/<w:delText> siblings fold into one
 * TextRun until interrupted by an inline child. <w:rPr>, activeComments,
 * trackedChange, and hyperlink apply to every TextRun produced. */
function readRun(
	view: DocView,
	node: XmlNode,
	activeComments: Set<string>,
	trackedChange: TrackedChange | undefined,
	hyperlink: Hyperlink | undefined,
	state: WalkState,
): Run[] {
	const runProperties = node.findChild("w:rPr");
	const out: Run[] = [];
	let pendingText = "";

	function flushText(): void {
		if (pendingText.length === 0) return;
		const run: TextRun = { type: "text", text: pendingText };
		if (runProperties) applyRunProperties(run, runProperties);
		if (activeComments.size > 0) run.comments = [...activeComments];
		if (trackedChange) run.trackedChange = trackedChange;
		if (hyperlink) run.hyperlink = hyperlink;
		out.push(run);
		pendingText = "";
	}

	for (const child of node.children) {
		if (child.tag === "w:rPr") continue;
		if (child.tag === "w:t" || child.tag === "w:delText") {
			pendingText += child.collectText();
			continue;
		}
		// Single-character text equivalents — fold into pendingText so they
		// share the surrounding rPr/tracking decoration. Each contributes
		// exactly one character to the AST text and to offset accounting.
		if (child.tag === "w:noBreakHyphen") {
			pendingText += "‑";
			continue;
		}
		if (child.tag === "w:softHyphen") {
			pendingText += "­";
			continue;
		}
		if (child.tag === "w:sym") {
			const font = child.getAttribute("w:font") ?? "";
			const charHex = child.getAttribute("w:char") ?? "";
			pendingText += decodeSym(font, charHex);
			continue;
		}
		if (child.tag === "w:drawing") {
			flushText();
			const drawing = readDrawing(view, child, state);
			if (drawing) {
				if (trackedChange && drawing.type === "image") {
					drawing.trackedChange = trackedChange;
				}
				out.push(drawing);
			}
			continue;
		}
		// Legacy embeds — surface as ChartRun placeholders so callers know
		// "something visual lives here." Underlying XML is preserved.
		if (child.tag === "w:pict" || child.tag === "w:object") {
			flushText();
			out.push({ type: "chart", kind: "drawing" });
			continue;
		}
		if (child.tag === "w:br" || child.tag === "w:cr") {
			flushText();
			// <w:cr> is an in-paragraph carriage return — semantically a line
			// break, same shape as <w:br w:type="line"/>.
			const kind =
				child.tag === "w:cr"
					? "line"
					: ((child.getAttribute("w:type") ?? "line") as
							| "page"
							| "line"
							| "column");
			out.push({ type: "break", kind });
			continue;
		}
		if (child.tag === "w:tab" || child.tag === "w:ptab") {
			flushText();
			out.push({ type: "tab" });
			continue;
		}
		if (child.tag === "w:footnoteReference") {
			flushText();
			const id = child.getAttribute("w:id");
			if (id)
				out.push({ type: "footnoteRef", kind: "footnote", id: `fn${id}` });
			continue;
		}
		if (child.tag === "w:endnoteReference") {
			flushText();
			const id = child.getAttribute("w:id");
			if (id) out.push({ type: "footnoteRef", kind: "endnote", id: `en${id}` });
		}
	}

	flushText();
	return out;
}

/** A <w:drawing> may wrap a picture (rendered as ImageRun) or a chart/shape/
 * SmartArt/etc. (rendered as ChartRun placeholder). */
function readDrawing(
	view: DocView,
	drawing: XmlNode,
	state: WalkState,
): ImageRun | ChartRun | null {
	const image = readImageFromDrawing(view, drawing, state);
	if (image) return image;
	if (drawing.findDescendant("c:chart"))
		return { type: "chart", kind: "chart" };
	if (drawing.findDescendant("dgm:relIds"))
		return { type: "chart", kind: "smartart" };
	if (drawing.findDescendant("wps:wsp"))
		return { type: "chart", kind: "shape" };
	return { type: "chart", kind: "drawing" };
}

/** Concatenate all <m:t> descendants in document order. The structural OOMath
 * (subs, sups, fractions) collapses to flat characters — degraded but readable. */
function collectMathText(node: XmlNode): string {
	let out = "";
	for (const child of node.children) {
		if (child.tag === "m:t") {
			out += child.collectText();
			continue;
		}
		out += collectMathText(child);
	}
	return out;
}

function applyRunProperties(run: TextRun, runProperties: XmlNode): void {
	const colorNode = runProperties.findChild("w:color");
	if (colorNode) {
		const value = colorNode.getAttribute("w:val");
		if (value && value !== "auto") run.color = value;
	}

	const highlightNode = runProperties.findChild("w:highlight");
	if (highlightNode) {
		const value = highlightNode.getAttribute("w:val");
		if (value && value !== "none") run.highlight = value;
	}

	if (runProperties.findChild("w:b")) run.bold = true;
	if (runProperties.findChild("w:i")) run.italic = true;

	const underlineNode = runProperties.findChild("w:u");
	if (underlineNode) {
		const value = underlineNode.getAttribute("w:val");
		if (value && value !== "none") run.underline = value;
	}

	if (runProperties.findChild("w:strike")) run.strike = true;

	const fontNode = runProperties.findChild("w:rFonts");
	if (fontNode) {
		const value =
			fontNode.getAttribute("w:ascii") ?? fontNode.getAttribute("w:hAnsi");
		if (value) run.font = value;
	}

	const sizeNode = runProperties.findChild("w:sz");
	if (sizeNode) {
		const value = sizeNode.getAttribute("w:val");
		if (value) run.sizeHalfPoints = Number(value);
	}

	const styleNode = runProperties.findChild("w:rStyle");
	if (styleNode) {
		const value = styleNode.getAttribute("w:val");
		// Skip FootnoteReference/EndnoteReference/CommentReference — those style
		// markers are part of the (foot|end)note/comment infrastructure and are
		// surfaced via the dedicated AST run types (footnoteRef, commentReference);
		// preserving them on the underlying TextRun would double-tag.
		if (
			value &&
			value !== "FootnoteReference" &&
			value !== "EndnoteReference" &&
			value !== "CommentReference"
		) {
			run.runStyle = value;
		}
	}
}

function readImageFromDrawing(
	view: DocView,
	drawing: XmlNode,
	state: WalkState,
): ImageRun | null {
	const blip = drawing.findDescendant("a:blip");
	if (!blip) return null;
	const relationshipId =
		blip.getAttribute("r:embed") ?? blip.getAttribute("r:link");
	if (!relationshipId) return null;
	const relationship = view.imagesByRelationshipId.get(relationshipId);
	if (!relationship) return null;

	const id = `img${state.imageIndex++}`;
	view.imageById.set(id, {
		relationshipId,
		partName: relationship.partName,
		contentType: relationship.contentType,
	});

	const docPr = drawing.findDescendant("wp:docPr");
	const alt = docPr
		? (docPr.getAttribute("descr") ?? docPr.getAttribute("name"))
		: undefined;

	const extent = drawing.findDescendant("wp:extent");
	const widthEmu = extent
		? Number(extent.getAttribute("cx") ?? "0")
		: undefined;
	const heightEmu = extent
		? Number(extent.getAttribute("cy") ?? "0")
		: undefined;

	return {
		type: "image",
		id,
		contentType: relationship.contentType,
		hash: "",
		widthEmu,
		heightEmu,
		alt: alt || undefined,
	};
}

function readTable(
	view: DocView,
	node: XmlNode,
	id: string,
	state: WalkState,
): Table {
	const grid = readTableGrid(node);
	const width = readTableWidth(node);
	// Register table-level revisions first, in tree order (tblPr before tblGrid,
	// both before the rows), so tcN ids match the apply.ts walk order.
	readTablePropertyRevision(view, node, id, state);
	readGridRevision(view, node, id, state);
	const rows: TableRow[] = [];
	let rowIndex = 0;
	for (const child of node.children) {
		if (child.tag !== "w:tr") continue;
		// Register the row-level revision before its cells so tcN ids match the
		// apply.ts walk order (row marker → per cell: cell marker → content).
		const rowChange = readRowRevision(view, child, id, state);
		const cells: TableCell[] = [];
		let columnIndex = 0;
		for (const cellNode of child.children) {
			if (cellNode.tag !== "w:tc") continue;
			cells.push(
				readTableCell(view, cellNode, id, rowIndex, columnIndex, state),
			);
			columnIndex++;
		}
		const row: TableRow = { cells };
		if (rowChange) row.trackedChange = rowChange;
		rows.push(row);
		rowIndex++;
	}
	const table: Table = { id, type: "table", grid, rows };
	if (width) table.width = width;
	return table;
}

function readTablePropertyRevision(
	view: DocView,
	table: XmlNode,
	tableId: string,
	state: WalkState,
): void {
	const tblPr = table.findChild("w:tblPr");
	const change = tblPr?.findChild("w:tblPrChange");
	if (!tblPr || !change) return;
	registerTableRevision(
		view,
		change,
		tblPr.children,
		"tblPrChange",
		tableId,
		state,
	);
}

function readGridRevision(
	view: DocView,
	table: XmlNode,
	tableId: string,
	state: WalkState,
): void {
	const tblGrid = table.findChild("w:tblGrid");
	const change = tblGrid?.findChild("w:tblGridChange");
	if (!tblGrid || !change) return;
	registerTableRevision(
		view,
		change,
		tblGrid.children,
		"tblGridChange",
		tableId,
		state,
	);
}

function readRowRevision(
	view: DocView,
	row: XmlNode,
	tableId: string,
	state: WalkState,
): TrackedChange | undefined {
	const trPr = row.findChild("w:trPr");
	const marker = trPr?.children.find(
		(child) => child.tag === "w:ins" || child.tag === "w:del",
	);
	if (!trPr || !marker) return undefined;
	return registerTableRevision(
		view,
		marker,
		trPr.children,
		marker.tag === "w:ins" ? "rowIns" : "rowDel",
		tableId,
		state,
	);
}

function readCellRevision(
	view: DocView,
	cell: XmlNode,
	tableId: string,
	state: WalkState,
): TrackedChange | undefined {
	const tcPr = cell.findChild("w:tcPr");
	const marker = tcPr?.children.find(
		(child) => child.tag === "w:cellIns" || child.tag === "w:cellDel",
	);
	if (!tcPr || !marker) return undefined;
	return registerTableRevision(
		view,
		marker,
		tcPr.children,
		marker.tag === "w:cellIns" ? "cellIns" : "cellDel",
		tableId,
		state,
	);
}

function registerTableRevision(
	view: DocView,
	marker: XmlNode,
	parent: XmlNode[],
	kind: TrackedChange["kind"],
	tableId: string,
	state: WalkState,
): TrackedChange {
	const trackedChangeId = `tc${state.trackedChangeIndex++}`;
	const change: TrackedChange = {
		id: trackedChangeId,
		kind,
		author: marker.getAttribute("w:author") ?? "",
		date: marker.getAttribute("w:date") ?? "",
		revisionId: marker.getAttribute("w:id") ?? "",
	};
	view.trackedChangeReferences.set(trackedChangeId, {
		node: marker,
		parent,
		blockId: tableId,
		kind,
	});
	return change;
}

function readTableGrid(table: XmlNode): number[] {
	const tblGrid = table.findChild("w:tblGrid");
	if (!tblGrid) return [];
	const widths: number[] = [];
	for (const col of tblGrid.findChildren("w:gridCol")) {
		const raw = col.getAttribute("w:w");
		const value = raw ? Number(raw) : NaN;
		widths.push(Number.isFinite(value) ? value : 0);
	}
	return widths;
}

function readTableWidth(table: XmlNode): TableWidth | undefined {
	const tblPr = table.findChild("w:tblPr");
	const tblW = tblPr?.findChild("w:tblW");
	if (!tblW) return undefined;
	return readWidth(tblW);
}

function readWidth(node: XmlNode): TableWidth | undefined {
	const raw = node.getAttribute("w:w");
	const unit = node.getAttribute("w:type");
	const value = raw ? Number(raw) : NaN;
	const resolvedUnit =
		unit === "dxa" || unit === "pct" || unit === "auto" || unit === "nil"
			? unit
			: "dxa";
	if (!Number.isFinite(value)) return undefined;
	return { value, unit: resolvedUnit };
}

function readTableCell(
	view: DocView,
	cellNode: XmlNode,
	tableId: string,
	rowIndex: number,
	columnIndex: number,
	state: WalkState,
): TableCell {
	// Register cell-level revisions before content so tcN ids match the apply.ts
	// walk order: the cellIns/cellDel marker first, then the tcPrChange.
	const cellChange = readCellRevision(view, cellNode, tableId, state);
	const cellPropertyChange = readCellPropertyRevision(
		view,
		cellNode,
		tableId,
		state,
	);
	const cell: TableCell = {
		blocks: readCellBlocks(
			view,
			cellNode,
			tableId,
			rowIndex,
			columnIndex,
			state,
		),
	};
	if (cellChange) cell.trackedChange = cellChange;
	else if (cellPropertyChange) cell.trackedChange = cellPropertyChange;
	const tcPr = cellNode.findChild("w:tcPr");
	if (!tcPr) return cell;
	const gridSpanNode = tcPr.findChild("w:gridSpan");
	if (gridSpanNode) {
		const raw = gridSpanNode.getAttribute("w:val");
		const value = raw ? Number(raw) : NaN;
		if (Number.isFinite(value) && value > 1) cell.gridSpan = value;
	}
	const vMergeNode = tcPr.findChild("w:vMerge");
	if (vMergeNode) {
		// Per ECMA-376 §17.4.84: w:val="restart" begins a new vertical merge;
		// a bare <w:vMerge/> (or w:val="continue") continues the merge from
		// the cell above.
		const raw = vMergeNode.getAttribute("w:val");
		cell.vMerge = raw === "restart" ? "restart" : "continue";
	}
	const tcW = tcPr.findChild("w:tcW");
	if (tcW) {
		const width = readWidth(tcW);
		if (width) cell.width = width;
	}
	return cell;
}

function readCellPropertyRevision(
	view: DocView,
	cell: XmlNode,
	tableId: string,
	state: WalkState,
): TrackedChange | undefined {
	const tcPr = cell.findChild("w:tcPr");
	const change = tcPr?.findChild("w:tcPrChange");
	if (!tcPr || !change) return undefined;
	return registerTableRevision(
		view,
		change,
		tcPr.children,
		"tcPrChange",
		tableId,
		state,
	);
}

function readCellBlocks(
	view: DocView,
	cell: XmlNode,
	tableId: string,
	rowIndex: number,
	columnIndex: number,
	state: WalkState,
): Block[] {
	const blocks: Block[] = [];
	let paragraphIndex = 0;
	let nestedTableIndex = 0;
	const cellPrefix = `${tableId}:r${rowIndex}c${columnIndex}`;
	for (const child of cell.children) {
		if (child.tag === "w:p") {
			const id = `${cellPrefix}:p${paragraphIndex++}`;
			blocks.push(readParagraph(view, child, id, state));
			view.blockReferences.set(id, { node: child, parent: cell.children });
			continue;
		}
		// Nested tables are legal inside a cell (Word emits them for compound
		// rubric layouts, etc.). Recurse with a chained id so locators like
		// `t0:r2c1:t0:r0c0:p0` resolve via the existing locator parser's
		// recursive `cell.inner`.
		if (child.tag === "w:tbl") {
			const id = `${cellPrefix}:t${nestedTableIndex++}`;
			blocks.push(readTable(view, child, id, state));
			view.blockReferences.set(id, { node: child, parent: cell.children });
		}
	}
	return blocks;
}

function readRelationships(view: DocView): void {
	const relationships = XmlNode.findRoot(
		view.relationshipsTree,
		"Relationships",
	);
	if (!relationships) return;
	for (const child of relationships.children) {
		if (child.tag !== "Relationship") continue;
		const type = child.getAttribute("Type");
		const relationshipId = child.getAttribute("Id");
		const target = child.getAttribute("Target");
		if (!relationshipId || !target) continue;
		if (type === RELATIONSHIP_NAMESPACE_IMAGE) {
			const partName = target.startsWith("/")
				? target.slice(1)
				: `word/${target}`;
			const contentType = lookupContentType(view, partName);
			view.imagesByRelationshipId.set(relationshipId, {
				partName,
				contentType,
			});
			continue;
		}
		if (type === RELATIONSHIP_NAMESPACE_HYPERLINK) {
			view.hyperlinksByRelationshipId.set(relationshipId, { url: target });
		}
	}
}

function lookupContentType(view: DocView, partName: string): string {
	const types = XmlNode.findRoot(view.contentTypesTree, "Types");
	if (!types) return "application/octet-stream";

	for (const child of types.children) {
		if (child.tag !== "Override") continue;
		if (child.getAttribute("PartName") === `/${partName}`) {
			return child.getAttribute("ContentType") ?? "application/octet-stream";
		}
	}

	const extension = partName.split(".").pop()?.toLowerCase() ?? "";
	for (const child of types.children) {
		if (child.tag !== "Default") continue;
		if (child.getAttribute("Extension")?.toLowerCase() === extension) {
			return child.getAttribute("ContentType") ?? "application/octet-stream";
		}
	}
	return "application/octet-stream";
}

function readDocProperties(view: DocView): DocProperties {
	if (!view.corePropertiesTree) return {};
	const root = XmlNode.findRoot(view.corePropertiesTree, "cp:coreProperties");
	if (!root) return {};
	const out: DocProperties = {};
	const title = root.findChild("dc:title");
	if (title) out.title = title.collectText();
	const author = root.findChild("dc:creator");
	if (author) out.author = author.collectText();
	const created = root.findChild("dcterms:created");
	if (created) out.created = created.collectText();
	const modified = root.findChild("dcterms:modified");
	if (modified) out.modified = modified.collectText();
	return out;
}

function readComments(
	view: DocView,
	anchors: Map<string, CommentAnchor>,
): Comment[] {
	if (!view.commentsTree) return [];
	const root = XmlNode.findRoot(view.commentsTree, "w:comments");
	if (!root) return [];

	const commentIdByParaId = new Map<string, string>();
	for (const child of root.children) {
		if (child.tag !== "w:comment") continue;
		const numericId = child.getAttribute("w:id");
		if (numericId == null) continue;
		const paragraph = child.findChild("w:p");
		const paraId = paragraph?.getAttribute("w14:paraId");
		if (paraId) commentIdByParaId.set(paraId, `c${numericId}`);
	}

	const extendedByParaId = readCommentsExtended(view);
	const comments: Comment[] = [];

	for (const child of root.children) {
		if (child.tag !== "w:comment") continue;
		const numericId = child.getAttribute("w:id");
		if (numericId == null) continue;
		const commentId = `c${numericId}`;
		const author = child.getAttribute("w:author") ?? "";
		const date = child.getAttribute("w:date") ?? "";
		const initials = child.getAttribute("w:initials");
		const text = child.collectText();
		const anchor = anchors.get(commentId) ?? {
			startBlockId: "",
			startOffset: 0,
			endBlockId: "",
			endOffset: 0,
		};

		const paragraph = child.findChild("w:p");
		const paraId = paragraph?.getAttribute("w14:paraId");
		const meta = paraId ? (extendedByParaId.get(paraId) ?? {}) : {};
		const parentCommentId = meta.parentParaId
			? commentIdByParaId.get(meta.parentParaId)
			: undefined;

		comments.push({
			id: commentId,
			author,
			...(initials ? { initials } : {}),
			date,
			text,
			anchor,
			...(parentCommentId ? { parentId: parentCommentId } : {}),
			...(meta.resolved !== undefined ? { resolved: meta.resolved } : {}),
		});
		view.commentReferences.set(commentId, {
			node: child,
			parent: root.children,
		});
	}
	return comments;
}

function readCommentsExtended(
	view: DocView,
): Map<string, { parentParaId?: string; resolved?: boolean }> {
	const out = new Map<string, { parentParaId?: string; resolved?: boolean }>();
	if (!view.commentsExtTree) return out;
	const root = XmlNode.findRoot(view.commentsExtTree, "w15:commentsEx");
	if (!root) return out;

	for (const child of root.children) {
		if (child.tag !== "w15:commentEx") continue;
		const paragraphId = child.getAttribute("w15:paraId");
		if (!paragraphId) continue;
		const resolvedAttribute = child.getAttribute("w15:done");
		const parentParagraphId = child.getAttribute("w15:paraIdParent");
		const entry: { parentParaId?: string; resolved?: boolean } = {};
		if (resolvedAttribute === "1") entry.resolved = true;
		else if (resolvedAttribute === "0") entry.resolved = false;
		if (parentParagraphId) entry.parentParaId = parentParagraphId;
		out.set(paragraphId, entry);
	}
	return out;
}

/** Read footnotes.xml or endnotes.xml. Skips Word's reserved separator/
 * continuationSeparator entries (w:type set, never referenced from the body). */
function readNotes(
	tree: XmlNode[] | undefined,
	rootTag: string,
	itemTag: string,
	idPrefix: "fn" | "en",
): Footnote[] {
	if (!tree) return [];
	const root = XmlNode.findRoot(tree, rootTag);
	if (!root) return [];
	const out: Footnote[] = [];
	for (const child of root.children) {
		if (child.tag !== itemTag) continue;
		if (child.getAttribute("w:type")) continue;
		const numericId = child.getAttribute("w:id");
		if (numericId == null) continue;
		const text = child.collectText().replace(/\s+/g, " ").trim();
		out.push({ id: `${idPrefix}${numericId}`, text });
	}
	return out;
}
