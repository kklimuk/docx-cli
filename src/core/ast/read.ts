import { XmlNode } from "../parser";
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
	Table,
	TableCell,
	TableRow,
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
			blocks.push({ id, type: "sectionBreak" });
			view.blockReferences.set(id, { node: child, parent: body.children });
		}
	}
	return blocks;
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
		applyParagraphProperties(paragraph, paragraphProperties);
	}

	const context: WalkContext = {
		view,
		blockId: id,
		paragraph,
		activeComments: new Set<string>(),
		state,
		offsetRef: { value: 0 },
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
};

function walkRunContainer(
	context: WalkContext,
	container: XmlNode,
	trackedChange: TrackedChange | undefined,
	hyperlink: Hyperlink | undefined,
): void {
	for (const child of container.children) {
		if (child.tag === "w:pPr") continue;

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
		paragraph.list = { level, numId: id };
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
			if (drawing) out.push(drawing);
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
	const rows: TableRow[] = [];
	let rowIndex = 0;
	for (const child of node.children) {
		if (child.tag !== "w:tr") continue;
		const cells: TableCell[] = [];
		let columnIndex = 0;
		for (const cellNode of child.children) {
			if (cellNode.tag !== "w:tc") continue;
			cells.push({
				blocks: readCellBlocks(
					view,
					cellNode,
					id,
					rowIndex,
					columnIndex,
					state,
				),
			});
			columnIndex++;
		}
		rows.push({ cells });
		rowIndex++;
	}
	return { id, type: "table", rows };
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
	for (const child of cell.children) {
		if (child.tag === "w:p") {
			const id = `${tableId}:r${rowIndex}c${columnIndex}:p${paragraphIndex++}`;
			blocks.push(readParagraph(view, child, id, state));
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
