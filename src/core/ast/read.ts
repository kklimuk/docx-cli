import { XmlNode } from "../parser";
import type { DocView } from "./doc-view";
import type {
	Block,
	Comment,
	CommentAnchor,
	Doc,
	DocProperties,
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

type WalkState = {
	imageIndex: number;
	commentAnchors: Map<string, CommentAnchor>;
	openComments: Map<string, { blockId: string; offset: number }>;
};

export function buildDoc(view: DocView, path: string): Doc {
	readImageRelationships(view);
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
		commentAnchors: new Map(),
		openComments: new Map(),
	};

	const blocks = readBlocks(view, body, state);
	const comments = readComments(view, state.commentAnchors);

	return { schemaVersion: 1, path, properties, blocks, comments };
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

	const activeComments = new Set<string>();
	let offset = 0;

	for (const child of node.children) {
		if (child.tag === "w:pPr") continue;

		if (child.tag === "w:commentRangeStart") {
			const commentId = child.getAttribute("w:id");
			if (commentId) {
				const key = `c${commentId}`;
				activeComments.add(key);
				state.openComments.set(key, { blockId: id, offset });
			}
			continue;
		}

		if (child.tag === "w:commentRangeEnd") {
			const commentId = child.getAttribute("w:id");
			if (commentId) {
				const key = `c${commentId}`;
				activeComments.delete(key);
				const opened = state.openComments.get(key);
				if (opened) {
					state.commentAnchors.set(key, {
						startBlockId: opened.blockId,
						startOffset: opened.offset,
						endBlockId: id,
						endOffset: offset,
					});
					state.openComments.delete(key);
				}
			}
			continue;
		}

		if (child.tag === "w:r") {
			const run = readRun(view, child, activeComments, undefined, state);
			if (run) {
				if (run.type === "text") offset += run.text.length;
				paragraph.runs.push(run);
			}
			continue;
		}

		if (child.tag === "w:ins" || child.tag === "w:del") {
			const change: TrackedChange = {
				kind: child.tag === "w:ins" ? "ins" : "del",
				author: child.getAttribute("w:author") ?? "",
				date: child.getAttribute("w:date") ?? "",
				revisionId: child.getAttribute("w:id") ?? "",
			};
			for (const inner of child.children) {
				if (inner.tag === "w:commentRangeStart") {
					const commentId = inner.getAttribute("w:id");
					if (commentId) {
						const key = `c${commentId}`;
						activeComments.add(key);
						state.openComments.set(key, { blockId: id, offset });
					}
					continue;
				}
				if (inner.tag === "w:commentRangeEnd") {
					const commentId = inner.getAttribute("w:id");
					if (commentId) {
						const key = `c${commentId}`;
						activeComments.delete(key);
						const opened = state.openComments.get(key);
						if (opened) {
							state.commentAnchors.set(key, {
								startBlockId: opened.blockId,
								startOffset: opened.offset,
								endBlockId: id,
								endOffset: offset,
							});
							state.openComments.delete(key);
						}
					}
					continue;
				}
				if (inner.tag !== "w:r") continue;
				const run = readRun(view, inner, activeComments, change, state);
				if (!run) continue;
				if (run.type === "text") offset += run.text.length;
				paragraph.runs.push(run);
			}
		}
	}

	return paragraph;
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

function readRun(
	view: DocView,
	node: XmlNode,
	activeComments: Set<string>,
	trackedChange: TrackedChange | undefined,
	state: WalkState,
): Run | null {
	const runProperties = node.findChild("w:rPr");

	for (const child of node.children) {
		if (child.tag === "w:drawing") {
			const image = readImageFromDrawing(view, child, state);
			if (image) return image;
		}
		if (child.tag === "w:br") {
			const kind = (child.getAttribute("w:type") ?? "line") as
				| "page"
				| "line"
				| "column";
			return { type: "break", kind };
		}
		if (child.tag === "w:tab") {
			return { type: "tab" };
		}
	}

	let combinedText = "";
	for (const child of node.children) {
		if (child.tag === "w:t") combinedText += child.collectText();
	}
	if (combinedText.length === 0) return null;

	const run: TextRun = { type: "text", text: combinedText };
	if (runProperties) applyRunProperties(run, runProperties);
	if (activeComments.size > 0) run.comments = [...activeComments];
	if (trackedChange) run.trackedChange = trackedChange;
	return run;
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

function readImageRelationships(view: DocView): void {
	const relationships = XmlNode.findRoot(
		view.relationshipsTree,
		"Relationships",
	);
	if (!relationships) return;
	for (const child of relationships.children) {
		if (child.tag !== "Relationship") continue;
		if (child.getAttribute("Type") !== RELATIONSHIP_NAMESPACE_IMAGE) continue;
		const relationshipId = child.getAttribute("Id");
		const target = child.getAttribute("Target");
		if (!relationshipId || !target) continue;
		const partName = target.startsWith("/")
			? target.slice(1)
			: `word/${target}`;
		const contentType = lookupContentType(view, partName);
		view.imagesByRelationshipId.set(relationshipId, { partName, contentType });
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
