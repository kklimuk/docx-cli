import { generateParaId } from "../ast/document/comments";
import type { FindView } from "../find";
import { w } from "../jsx";
import {
	isRunBearingWrapper,
	runTextLength,
	sliceRun,
	sumRunBearingTextLength,
	XmlNode,
} from "../parser";

export { generateParaId };

/** Whether a run-bearing wrapper's contents are visible in the chosen
 *  view. Mirrors `isWrapperVisibleInView` in `cli/replace/replace-span.tsx`
 *  and `isRunVisibleInView` in `core/find/index.ts` — they MUST agree, or
 *  `find → comments add` (and `find → replace`) misalign. */
function isWrapperVisibleInView(tag: string, view: FindView): boolean {
	if (!isRunBearingWrapper(tag)) return false;
	if (view === "current") return true;
	if (view === "accepted") return tag !== "w:del" && tag !== "w:moveFrom";
	return tag !== "w:ins" && tag !== "w:moveTo";
}

/** Sum the text length of `children` counting only runs visible in the
 *  chosen view. Used for paragraph length / span bounds checks when
 *  placing comment markers. */
function sumVisibleTextLength(children: XmlNode[], view: FindView): number {
	let total = 0;
	for (const child of children) {
		if (child.tag === "w:r") {
			total += runTextLength(child);
			continue;
		}
		if (isWrapperVisibleInView(child.tag, view)) {
			total += sumVisibleTextLength(child.children, view);
		}
	}
	return total;
}

export function authorInitials(author: string): string {
	const parts = author.trim().split(/\s+/).filter(Boolean);
	if (parts.length === 0) return "?";
	const first = parts[0]?.charAt(0) ?? "";
	const last =
		parts.length > 1 ? (parts[parts.length - 1]?.charAt(0) ?? "") : "";
	return (first + last).toUpperCase() || "?";
}

export type CommentSpan = { start: number; end: number };

export class SpanOutOfRangeError extends Error {
	constructor(message: string) {
		super(message);
		this.name = "SpanOutOfRangeError";
	}
}

export function paragraphTextLength(
	paragraph: XmlNode,
	view: FindView = "current",
): number {
	if (view === "current") return sumRunBearingTextLength(paragraph.children);
	return sumVisibleTextLength(paragraph.children, view);
}

export function addCommentMarkersToParagraph(
	paragraph: XmlNode,
	commentId: string,
	span?: CommentSpan,
	view: FindView = "current",
): void {
	const total = paragraphTextLength(paragraph, view);
	const range: CommentSpan = span ?? { start: 0, end: total };
	if (range.start < 0 || range.end > total || range.start > range.end) {
		throw new SpanOutOfRangeError(
			`Span ${range.start}-${range.end} out of paragraph length ${total}`,
		);
	}
	placeMarkersInParagraph(
		paragraph,
		[
			{ offset: range.start, node: commentRangeStartMarker(commentId) },
			{
				offset: range.end,
				node: commentRangeEndMarker(commentId),
				follower: commentReferenceRun(commentId),
			},
		],
		view,
	);
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
	view: FindView = "current",
): void {
	if (startParagraph === endParagraph) {
		addCommentMarkersToParagraph(
			startParagraph,
			commentId,
			{
				start: startOffset,
				end: endOffset,
			},
			view,
		);
		return;
	}
	placeMarkersInParagraph(
		startParagraph,
		[{ offset: startOffset, node: commentRangeStartMarker(commentId) }],
		view,
	);
	placeMarkersInParagraph(
		endParagraph,
		[
			{
				offset: endOffset,
				node: commentRangeEndMarker(commentId),
				follower: commentReferenceRun(commentId),
			},
		],
		view,
	);
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

function commentRangeStartMarker(commentId: string): XmlNode {
	return <w.commentRangeStart w-id={commentId} />;
}

function commentRangeEndMarker(commentId: string): XmlNode {
	return <w.commentRangeEnd w-id={commentId} />;
}

function commentReferenceRun(commentId: string): XmlNode {
	return (
		<w.r>
			<w.rPr>
				<w.rStyle w-val="CommentReference" />
			</w.rPr>
			<w.commentReference w-id={commentId} />
		</w.r>
	);
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
	view: FindView = "current",
): void {
	if (markers.length === 0) return;
	const total = paragraphTextLength(paragraph, view);
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

	paragraph.children = walkAndPlace(paragraph.children, pending, state, view);
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
	view: FindView,
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
			// Wrapper invisible in the chosen view: pass through with no
			// offset change. Don't flush pending markers — they belong to a
			// boundary in the visible-text coordinate space, not to this
			// invisible wrapper. The next visible run will pick them up.
			if (!isWrapperVisibleInView(child.tag, view)) {
				result.push(child);
				continue;
			}
			// Boundary at the visible wrapper's start: place markers BEFORE
			// descending so they sit outside the wrapper when offsets align.
			flushAtCurrentOffset(result, pending, state);

			const innerChildren = walkAndPlace(child.children, pending, state, view);
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

export type ParagraphCommentMarker = { id: string; kind: "start" | "end" };

/** Snapshot the comment range markers (`<w:commentRangeStart>` /
 *  `<w:commentRangeEnd>`) and reference runs that live directly inside
 *  `paragraph`, **removing them** from the paragraph. Used by `edit` to lift a
 *  paragraph's comment anchors out of the way before its content is rebuilt,
 *  so they can be re-placed afterward instead of collapsing to a zero-length
 *  range (the bug where editing a commented paragraph orphaned the comment).
 *  Returns one entry per marker (a fully-enclosed comment yields a `start` and
 *  an `end`; a cross-paragraph half yields just one). */
export function extractCommentMarkers(
	paragraph: XmlNode,
): ParagraphCommentMarker[] {
	const found: ParagraphCommentMarker[] = [];
	function walk(node: XmlNode): void {
		const kept: XmlNode[] = [];
		for (const child of node.children) {
			if (child.tag === "w:commentRangeStart") {
				const id = child.getAttribute("w:id");
				if (id) found.push({ id, kind: "start" });
				continue;
			}
			if (child.tag === "w:commentRangeEnd") {
				const id = child.getAttribute("w:id");
				if (id) found.push({ id, kind: "end" });
				continue;
			}
			if (child.tag === "w:r" && runHasCommentReference(child)) continue;
			if (isRunBearingWrapper(child.tag)) walk(child);
			kept.push(child);
		}
		node.children = kept;
	}
	walk(paragraph);
	return found;
}

function runHasCommentReference(run: XmlNode): boolean {
	return run.children.some((child) => child.tag === "w:commentReference");
}

/** Re-place comment markers snapshotted by {@link extractCommentMarkers} so they
 *  bracket `paragraph`'s (rebuilt) content: each `start` re-anchors to the
 *  paragraph start, each `end` to the paragraph end — so a comment that was on
 *  the paragraph now spans the new content rather than collapsing. Returns the
 *  ids that could NOT be re-anchored because the new paragraph has no text
 *  (caller should mark those comments resolved). Offsets use the `current`
 *  view so the range covers all runs regardless of tracked-change wrappers. */
export function reanchorCommentMarkers(
	paragraph: XmlNode,
	markers: ParagraphCommentMarker[],
	view: FindView = "current",
): string[] {
	if (markers.length === 0) return [];
	const total = paragraphTextLength(paragraph, view);
	if (total === 0) return [...new Set(markers.map((marker) => marker.id))];
	const specs: MarkerSpec[] = markers.map((marker) =>
		marker.kind === "start"
			? { offset: 0, node: commentRangeStartMarker(marker.id) }
			: {
					offset: total,
					node: commentRangeEndMarker(marker.id),
					follower: commentReferenceRun(marker.id),
				},
	);
	placeMarkersInParagraph(paragraph, specs, view);
	return [];
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
