import type { DocView } from "@core";
import { w, w15 } from "@core/jsx";
import { registerPart } from "@core/package";
import {
	isRunBearingWrapper,
	runTextLength,
	sliceRun,
	sumRunBearingTextLength,
	XmlNode,
} from "@core/parser";

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
	return sumRunBearingTextLength(paragraph.children);
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

/**
 * Wrap `target` in `commentRangeStart` / `commentRangeEnd` siblings (plus the
 * commentReference run). Used when the target run has zero text length (e.g.
 * a run containing only a `<w:drawing>`), where offset-based placement would
 * collapse to an empty span.
 */
export function addCommentMarkersAroundRun(
	paragraph: XmlNode,
	target: XmlNode,
	commentId: string,
): boolean {
	function walk(parent: XmlNode): boolean {
		const children = parent.children;
		for (let index = 0; index < children.length; index++) {
			const child = children[index];
			if (child === target) {
				children.splice(
					index,
					1,
					commentRangeStartMarker(commentId),
					target,
					commentRangeEndMarker(commentId),
					commentReferenceRun(commentId),
				);
				return true;
			}
			if (child && isRunBearingWrapper(child.tag)) {
				if (walk(child)) return true;
			}
		}
		return false;
	}
	return walk(paragraph);
}

/**
 * Returns the text-offset range that `target` occupies inside `paragraph`,
 * or null if `target` is not a descendant of the paragraph. Recurses through
 * every run-bearing wrapper (see `RUN_BEARING_WRAPPER_TAGS`).
 */
export function findElementOffsetsInParagraph(
	paragraph: XmlNode,
	target: XmlNode,
): { start: number; end: number } | null {
	let cursor = 0;
	let result: { start: number; end: number } | null = null;
	function walk(children: XmlNode[]): boolean {
		for (const child of children) {
			if (child === target) {
				const start = cursor;
				cursor += sumRunBearingTextLength([child]);
				result = { start, end: cursor };
				return true;
			}
			if (child.tag === "w:r") {
				cursor += runTextLength(child);
				continue;
			}
			if (isRunBearingWrapper(child.tag)) {
				if (walk(child.children)) return true;
			}
		}
		return false;
	}
	walk(paragraph.children);
	return result;
}

/**
 * Walk `documentTree` to find the enclosing `<w:p>` of `target`, or null if
 * not found. Used when we have a node reference (hyperlink, drawing run) and
 * need its containing paragraph for comment-marker placement.
 */
export function findContainingParagraph(
	documentTree: XmlNode[],
	target: XmlNode,
): XmlNode | null {
	function walk(node: XmlNode): XmlNode | null {
		if (node.tag === "w:p" && containsNode(node, target)) return node;
		for (const child of node.children) {
			const found = walk(child);
			if (found) return found;
		}
		return null;
	}
	for (const root of documentTree) {
		const found = walk(root);
		if (found) return found;
	}
	return null;
}

function containsNode(haystack: XmlNode, needle: XmlNode): boolean {
	if (haystack === needle) return true;
	for (const child of haystack.children) {
		if (containsNode(child, needle)) return true;
	}
	return false;
}

export type AuditCommentAnchor =
	| { kind: "span"; paragraph: XmlNode; span: CommentSpan }
	| { kind: "run"; paragraph: XmlNode; run: XmlNode };

/**
 * Emit a comment for operations that OOXML can't track natively (hyperlinks
 * and image replacement). Anchored either to a text span (default) or around
 * a specific run (used for image runs where text length is zero). Returns
 * the numeric comment id.
 */
export function emitAuditComment(
	view: DocView,
	anchor: AuditCommentAnchor,
	options: { body: string; author: string; date: string },
): string {
	const numericId = nextCommentId(view);
	const paraId = generateParaId();

	if (anchor.kind === "span") {
		addCommentMarkersToParagraph(anchor.paragraph, numericId, anchor.span);
	} else {
		addCommentMarkersAroundRun(anchor.paragraph, anchor.run, numericId);
	}

	const commentsRoot = ensureCommentsPart(view);
	commentsRoot.children.push(
		<CommentBody
			options={{
				id: numericId,
				author: options.author,
				date: options.date,
				initials: authorInitials(options.author),
				paraId,
				text: options.body,
			}}
		/>,
	);

	return numericId;
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

	paragraph.children = walkAndPlace(paragraph.children, pending, state);
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

		if (isRunBearingWrapper(child.tag)) {
			// Boundary at the wrapper's start: place markers BEFORE descending so
			// they sit outside the wrapper when offsets align.
			flushAtCurrentOffset(result, pending, state);

			const innerChildren = walkAndPlace(child.children, pending, state);
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
