import type { DocView } from "@core";
import { w, w15 } from "@core/jsx";
import { registerPart } from "@core/package";
import { runTextLength, sliceRun, XmlNode } from "@core/parser";

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
	registerPart(view.relationshipsTree, view.contentTypesTree, {
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
	registerPart(view.relationshipsTree, view.contentTypesTree, {
		partName: "word/commentsExtended.xml",
		contentType: COMMENTS_EXT_CONTENT_TYPE,
		relationshipType: COMMENTS_EXT_REL_TYPE,
		target: "commentsExtended.xml",
	});
	return root;
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
	placeMarkersInParagraph(paragraph, [
		{ offset: range.start, node: commentRangeStartMarker(commentId) },
		{
			offset: range.end,
			node: commentRangeEndMarker(commentId),
			follower: commentReferenceRun(commentId),
		},
	]);
}

/**
 * Drop the start marker at `startOffset` of `startParagraph` and the end
 * marker (plus reference run) at `endOffset` of `endParagraph`. When both
 * paragraphs are the same node, behaves identically to
 * `addCommentMarkersToParagraph` with that span. Intermediate paragraphs
 * between the two are not touched — Word treats any text between matching
 * `<w:commentRangeStart>` / `<w:commentRangeEnd>` as covered.
 */
export function addCommentRangeMarkers(
	startParagraph: XmlNode,
	startOffset: number,
	endParagraph: XmlNode,
	endOffset: number,
	commentId: string,
): void {
	if (startParagraph === endParagraph) {
		addCommentMarkersToParagraph(startParagraph, commentId, {
			start: startOffset,
			end: endOffset,
		});
		return;
	}
	placeMarkersInParagraph(startParagraph, [
		{ offset: startOffset, node: commentRangeStartMarker(commentId) },
	]);
	placeMarkersInParagraph(endParagraph, [
		{
			offset: endOffset,
			node: commentRangeEndMarker(commentId),
			follower: commentReferenceRun(commentId),
		},
	]);
}

function commentRangeStartMarker(commentId: string): XmlNode {
	return (<w.commentRangeStart w-id={commentId} />) as XmlNode;
}

function commentRangeEndMarker(commentId: string): XmlNode {
	return (<w.commentRangeEnd w-id={commentId} />) as XmlNode;
}

function commentReferenceRun(commentId: string): XmlNode {
	return (
		<w.r>
			<w.rPr>
				<w.rStyle w-val="CommentReference" />
			</w.rPr>
			<w.commentReference w-id={commentId} />
		</w.r>
	) as XmlNode;
}

type MarkerSpec = {
	offset: number;
	node: XmlNode;
	follower?: XmlNode;
};

type PlacementState = {
	offset: number;
	placedCount: number;
};

function placeMarkersInParagraph(
	paragraph: XmlNode,
	markers: MarkerSpec[],
): void {
	if (markers.length === 0) return;
	const total = paragraphTextLength(paragraph);
	for (const marker of markers) {
		if (marker.offset < 0 || marker.offset > total) {
			throw new SpanOutOfRangeError(
				`Marker offset ${marker.offset} out of paragraph length ${total}`,
			);
		}
	}
	// Insertion order is the tiebreaker for markers at the same offset, so
	// callers can place start before end at a zero-length span.
	const pending: (MarkerSpec | null)[] = markers.slice();
	const state: PlacementState = { offset: 0, placedCount: 0 };

	paragraph.children = walkAndPlace(paragraph.children, pending, true, state);
	flushAtCurrentOffset(paragraph.children, pending, state);

	if (state.placedCount !== markers.length) {
		throw new SpanOutOfRangeError(
			`Could not place comment markers (placed ${state.placedCount} of ${markers.length})`,
		);
	}
}

function walkAndPlace(
	children: XmlNode[],
	pending: (MarkerSpec | null)[],
	isParagraphLevel: boolean,
	state: PlacementState,
): XmlNode[] {
	const result: XmlNode[] = [];
	for (const child of children) {
		if (child.tag === "w:r") {
			const length = runTextLength(child);
			const runStart = state.offset;
			const runEnd = state.offset + length;

			const splits: { at: number; index: number }[] = [];
			for (let i = 0; i < pending.length; i++) {
				const marker = pending[i];
				if (!marker) continue;
				if (runStart <= marker.offset && marker.offset <= runEnd) {
					splits.push({ at: marker.offset - runStart, index: i });
				}
			}
			splits.sort(
				(left, right) => left.at - right.at || left.index - right.index,
			);

			if (splits.length === 0) {
				result.push(child);
			} else {
				let cursor = 0;
				for (const split of splits) {
					if (split.at > cursor) {
						result.push(sliceRun(child, cursor, split.at));
					}
					const marker = pending[split.index];
					if (!marker) continue;
					result.push(marker.node);
					if (marker.follower) result.push(marker.follower);
					pending[split.index] = null;
					state.placedCount++;
					cursor = split.at;
				}
				if (cursor < length) {
					result.push(sliceRun(child, cursor, length));
				}
			}

			state.offset = runEnd;
			continue;
		}

		if (isParagraphLevel && (child.tag === "w:ins" || child.tag === "w:del")) {
			// Boundary at the wrapper's start: place markers BEFORE descending so
			// they sit outside the tracked-change wrapper when offsets align.
			flushAtCurrentOffset(result, pending, state);

			const innerChildren = walkAndPlace(child.children, pending, false, state);
			const wrapper = new XmlNode(child.tag, { ...child.attributes });
			wrapper.children = innerChildren;
			result.push(wrapper);
			continue;
		}

		if (child.tag === "w:pPr") {
			// pPr always comes first in a paragraph; markers must sit AFTER it.
			result.push(child);
			flushAtCurrentOffset(result, pending, state);
			continue;
		}

		// Non-run passthrough — markers BEFORE.
		flushAtCurrentOffset(result, pending, state);
		result.push(child);
	}
	return result;
}

function flushAtCurrentOffset(
	out: XmlNode[],
	pending: (MarkerSpec | null)[],
	state: PlacementState,
): void {
	for (let i = 0; i < pending.length; i++) {
		const marker = pending[i];
		if (!marker) continue;
		if (marker.offset !== state.offset) continue;
		out.push(marker.node);
		if (marker.follower) out.push(marker.follower);
		pending[i] = null;
		state.placedCount++;
	}
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
