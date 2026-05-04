import type { DocView } from "@core";
import { w, w15 } from "@core/jsx";
import { XmlNode } from "@core/parser";

const COMMENTS_REL_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const COMMENTS_EXT_REL_TYPE =
	"http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
const COMMENTS_CONTENT_TYPE =
	"application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const COMMENTS_EXT_CONTENT_TYPE =
	"application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const NS_W15 = "http://schemas.microsoft.com/office/word/2012/wordml";

export function nextCommentId(view: DocView): string {
	if (!view.commentsTree) return "0";
	const root = XmlNode.findRoot(view.commentsTree, "w:comments");
	if (!root) return "0";
	let highest = -1;
	for (const child of root.children) {
		if (child.tag !== "w:comment") continue;
		const idAttribute = child.getAttribute("w:id");
		if (idAttribute == null) continue;
		const numeric = Number(idAttribute);
		if (Number.isFinite(numeric) && numeric > highest) highest = numeric;
	}
	return String(highest + 1);
}

export function generateParaId(): string {
	const bytes = new Uint8Array(4);
	crypto.getRandomValues(bytes);
	let hex = "";
	for (const byte of bytes) hex += byte.toString(16).padStart(2, "0");
	return hex.toUpperCase();
}

export function authorInitials(author: string): string {
	const parts = author.trim().split(/\s+/).filter(Boolean);
	if (parts.length === 0) return "?";
	const first = parts[0]?.charAt(0) ?? "";
	const last =
		parts.length > 1 ? (parts[parts.length - 1]?.charAt(0) ?? "") : "";
	return (first + last).toUpperCase() || "?";
}

export function ensureCommentsPart(view: DocView): XmlNode {
	if (view.commentsTree) {
		const existing = XmlNode.findRoot(view.commentsTree, "w:comments");
		if (existing) return existing;
	}
	const root = <w.comments {...{ "xmlns:w": NS_W, "xmlns:w14": NS_W14 }} />;
	view.commentsTree = [root];
	registerPart(view, {
		partName: "word/comments.xml",
		contentType: COMMENTS_CONTENT_TYPE,
		relationshipType: COMMENTS_REL_TYPE,
		target: "comments.xml",
	});
	return root;
}

export function ensureCommentsExtPart(view: DocView): XmlNode {
	if (view.commentsExtTree) {
		const existing = XmlNode.findRoot(view.commentsExtTree, "w15:commentsEx");
		if (existing) return existing;
	}
	const root = (
		<w15.commentsEx
			{...{
				"xmlns:w15": NS_W15,
				"xmlns:mc":
					"http://schemas.openxmlformats.org/markup-compatibility/2006",
				"mc:Ignorable": "w15",
			}}
		/>
	);
	view.commentsExtTree = [root];
	registerPart(view, {
		partName: "word/commentsExtended.xml",
		contentType: COMMENTS_EXT_CONTENT_TYPE,
		relationshipType: COMMENTS_EXT_REL_TYPE,
		target: "commentsExtended.xml",
	});
	return root;
}

type PartRegistration = {
	partName: string;
	contentType: string;
	relationshipType: string;
	target: string;
};

function registerPart(view: DocView, part: PartRegistration): void {
	const relationships = XmlNode.findRoot(
		view.relationshipsTree,
		"Relationships",
	);
	if (relationships) {
		const alreadyLinked = relationships.children.some(
			(child) =>
				child.tag === "Relationship" &&
				child.getAttribute("Type") === part.relationshipType,
		);
		if (!alreadyLinked) {
			relationships.children.push(
				new XmlNode("Relationship", {
					Id: nextRelationshipId(relationships),
					Type: part.relationshipType,
					Target: part.target,
				}),
			);
		}
	}

	const types = XmlNode.findRoot(view.contentTypesTree, "Types");
	if (types) {
		const overrideExists = types.children.some(
			(child) =>
				child.tag === "Override" &&
				child.getAttribute("PartName") === `/${part.partName}`,
		);
		if (!overrideExists) {
			types.children.push(
				new XmlNode("Override", {
					PartName: `/${part.partName}`,
					ContentType: part.contentType,
				}),
			);
		}
	}
}

function nextRelationshipId(relationships: XmlNode): string {
	let highest = 0;
	for (const child of relationships.children) {
		if (child.tag !== "Relationship") continue;
		const id = child.getAttribute("Id");
		if (!id) continue;
		const match = id.match(/^rId(\d+)$/);
		if (!match) continue;
		const numeric = Number(match[1]);
		if (Number.isFinite(numeric) && numeric > highest) highest = numeric;
	}
	return `rId${highest + 1}`;
}

export type CommentSpan = { start: number; end: number };

export class SpanOutOfRangeError extends Error {
	constructor(message: string) {
		super(message);
		this.name = "SpanOutOfRangeError";
	}
}

export function paragraphTextLength(paragraph: XmlNode): number {
	let total = 0;
	for (const child of paragraph.children) {
		if (child.tag === "w:r") total += runTextLength(child);
		else if (child.tag === "w:ins" || child.tag === "w:del") {
			for (const inner of child.children) {
				if (inner.tag === "w:r") total += runTextLength(inner);
			}
		}
	}
	return total;
}

export function addCommentMarkersToParagraph(
	paragraph: XmlNode,
	commentId: string,
	span?: CommentSpan,
): void {
	const total = paragraphTextLength(paragraph);
	const range: CommentSpan = span ?? { start: 0, end: total };
	if (range.start < 0 || range.end > total || range.start > range.end) {
		throw new SpanOutOfRangeError(
			`Span ${range.start}-${range.end} out of paragraph length ${total}`,
		);
	}

	const newChildren: XmlNode[] = [];
	let offset = 0;
	let placedStart = false;
	let placedEnd = false;

	const startMarker = <w.commentRangeStart w-id={commentId} />;
	const endMarker = <w.commentRangeEnd w-id={commentId} />;
	const referenceRun = (
		<w.r>
			<w.rPr>
				<w.rStyle w-val="CommentReference" />
			</w.rPr>
			<w.commentReference w-id={commentId} />
		</w.r>
	);

	for (const child of paragraph.children) {
		if (child.tag !== "w:r") {
			// Non-run children (pPr, ins, del, etc.): pass through, but check
			// whether a boundary at offset==current should land before it.
			if (!placedStart && offset === range.start) {
				placedStart = true;
				if (child.tag === "w:pPr") {
					// Markers must come after pPr — push pPr first, then start.
					newChildren.push(child);
					newChildren.push(startMarker);
					continue;
				}
				newChildren.push(startMarker);
			}
			if (!placedEnd && offset === range.end && placedStart) {
				placedEnd = true;
				newChildren.push(endMarker);
				newChildren.push(referenceRun);
			}
			newChildren.push(child);
			continue;
		}

		const length = runTextLength(child);
		const runStart = offset;
		const runEnd = offset + length;

		// Collect splits inside this run, in order.
		const splits: { at: number; node: XmlNode }[] = [];
		if (!placedStart && runStart <= range.start && range.start <= runEnd) {
			splits.push({ at: range.start - runStart, node: startMarker });
			placedStart = true;
		}
		if (!placedEnd && runStart <= range.end && range.end <= runEnd) {
			splits.push({ at: range.end - runStart, node: endMarker });
			placedEnd = true;
		}
		splits.sort((leftSplit, rightSplit) => leftSplit.at - rightSplit.at);

		if (splits.length === 0) {
			newChildren.push(child);
		} else {
			let cursor = 0;
			for (const split of splits) {
				if (split.at > cursor) {
					newChildren.push(sliceRun(child, cursor, split.at));
				}
				newChildren.push(split.node);
				cursor = split.at;
			}
			if (cursor < length) {
				newChildren.push(sliceRun(child, cursor, length));
			}
			// If end marker was placed inside this run, reference run goes right after.
			if (splits.some((split) => split.node === endMarker)) {
				newChildren.push(referenceRun);
			}
		}

		offset = runEnd;
	}

	if (!placedStart && offset === range.start) {
		newChildren.push(startMarker);
		placedStart = true;
	}
	if (!placedEnd && offset === range.end) {
		newChildren.push(endMarker);
		newChildren.push(referenceRun);
		placedEnd = true;
	}

	if (!placedStart || !placedEnd) {
		throw new SpanOutOfRangeError(
			`Could not place comment markers (start placed: ${placedStart}, end placed: ${placedEnd})`,
		);
	}

	paragraph.children = newChildren;
}

function runTextLength(run: XmlNode): number {
	let total = 0;
	for (const child of run.children) {
		if (child.tag === "w:t") total += child.collectText().length;
	}
	return total;
}

function sliceRun(run: XmlNode, start: number, end: number): XmlNode {
	const sliced = new XmlNode("w:r", { ...run.attributes });
	let consumed = 0;
	for (const child of run.children) {
		if (child.tag === "w:t") {
			const text = child.collectText();
			const localStart = Math.max(0, start - consumed);
			const localEnd = Math.min(text.length, end - consumed);
			if (localStart < localEnd) {
				const slicedText = new XmlNode("w:t", { "xml:space": "preserve" });
				slicedText.children.push(
					XmlNode.textNode(text.slice(localStart, localEnd)),
				);
				sliced.children.push(slicedText);
			}
			consumed += text.length;
			continue;
		}
		// Non-text run children (rPr, etc.) — clone for safety.
		sliced.children.push(deepCloneNode(child));
	}
	return sliced;
}

function deepCloneNode(node: XmlNode): XmlNode {
	const clone = new XmlNode(node.tag, { ...node.attributes });
	if (node.text !== undefined) clone.text = node.text;
	for (const child of node.children) {
		clone.children.push(deepCloneNode(child));
	}
	return clone;
}

export type CommentBodyOptions = {
	id: string;
	author: string;
	date: string;
	initials: string;
	paraId: string;
	text: string;
};

export function CommentBody({
	options,
}: {
	options: CommentBodyOptions;
}): XmlNode {
	return (
		<w.comment
			w-id={options.id}
			w-author={options.author}
			w-date={options.date}
			w-initials={options.initials}
		>
			<w.p {...{ "w14:paraId": options.paraId, "w14:textId": "00000000" }}>
				<w.r>
					<w.t {...{ "xml:space": "preserve" }}>{options.text}</w.t>
				</w.r>
			</w.p>
		</w.comment>
	);
}

export function findCommentParaId(
	view: DocView,
	commentId: string,
): string | undefined {
	if (!view.commentsTree) return undefined;
	const root = XmlNode.findRoot(view.commentsTree, "w:comments");
	if (!root) return undefined;
	const numericId = commentId.startsWith("c") ? commentId.slice(1) : commentId;
	for (const child of root.children) {
		if (child.tag !== "w:comment") continue;
		if (child.getAttribute("w:id") !== numericId) continue;
		const paragraph = child.findChild("w:p");
		if (!paragraph) return undefined;
		return paragraph.getAttribute("w14:paraId");
	}
	return undefined;
}

export function ensureCommentParaId(
	view: DocView,
	commentId: string,
): string | undefined {
	if (!view.commentsTree) return undefined;
	const root = XmlNode.findRoot(view.commentsTree, "w:comments");
	if (!root) return undefined;
	const numericId = commentId.startsWith("c") ? commentId.slice(1) : commentId;
	for (const child of root.children) {
		if (child.tag !== "w:comment") continue;
		if (child.getAttribute("w:id") !== numericId) continue;
		const paragraph = child.findChild("w:p");
		if (!paragraph) return undefined;
		const existing = paragraph.getAttribute("w14:paraId");
		if (existing) return existing;
		const fresh = generateParaId();
		paragraph.setAttribute("w14:paraId", fresh);
		if (!root.attributes["xmlns:w14"]) {
			root.setAttribute("xmlns:w14", NS_W14);
		}
		return fresh;
	}
	return undefined;
}

export function findCommentByNumericId(
	view: DocView,
	numericId: string,
): { node: XmlNode; parent: XmlNode[] } | undefined {
	if (!view.commentsTree) return undefined;
	const root = XmlNode.findRoot(view.commentsTree, "w:comments");
	if (!root) return undefined;
	for (const child of root.children) {
		if (child.tag === "w:comment" && child.getAttribute("w:id") === numericId) {
			return { node: child, parent: root.children };
		}
	}
	return undefined;
}

export function removeCommentMarkers(
	documentTree: XmlNode[],
	numericId: string,
): void {
	const document = XmlNode.findRoot(documentTree, "w:document");
	if (!document) return;
	walkAndPruneCommentReferences(document, numericId);
}

function walkAndPruneCommentReferences(node: XmlNode, numericId: string): void {
	const filtered: XmlNode[] = [];
	for (const child of node.children) {
		if (
			(child.tag === "w:commentRangeStart" ||
				child.tag === "w:commentRangeEnd") &&
			child.getAttribute("w:id") === numericId
		) {
			continue;
		}
		if (child.tag === "w:r" && containsCommentReference(child, numericId)) {
			continue;
		}
		walkAndPruneCommentReferences(child, numericId);
		filtered.push(child);
	}
	node.children = filtered;
}

function containsCommentReference(run: XmlNode, numericId: string): boolean {
	for (const child of run.children) {
		if (
			child.tag === "w:commentReference" &&
			child.getAttribute("w:id") === numericId
		) {
			return true;
		}
	}
	return false;
}
