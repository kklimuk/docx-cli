import { runTextLength, sliceRun, wrapperContent, XmlNode } from "../parser";
import { convertTextToDelText, type TrackedMeta } from "../track-changes";
import { Del, Ins } from "../track-changes/emit";
import type { FindView } from "./index";
import {
	isWrapperVisibleInView,
	type ReplacementFormatting,
	replacementRuns,
	type Span,
	sumVisibleTextLength,
	type TrackedReplaceOptions,
} from "./replace-span";

/**
 * Tracked replace of a span that lies WHOLLY inside another revision's
 * `<w:ins>` / `<w:moveTo>` (issue #13). `replaceSpanInParagraph` dispatches
 * here only when every overlapping run shares that wrapper as its immediate
 * parent, the wrapper is a direct child of the paragraph, the span is
 * non-empty, and the replace is tracked — every other input stays on the
 * general walkers.
 *
 * The wrapper is split the way Word does it: it keeps its id and the text
 * before the span plus the cut, nested in the editor's `<w:del>` (deleting
 * inserted text is a del-of-ins); the editor's own `<w:ins>` carries the
 * replacement at paragraph level; the content after the span rides a copy of
 * the wrapper with a FRESH id — unless that tail is nothing but markers (a
 * comment range end and its reference run, a bookmark end), which stay in
 * the wrapper after the cut rather than becoming an empty revision. The one
 * exception to splitting is the editor's OWN `<w:ins>`, rebuilt in place
 * (same id, same date): the replacement is just more of that same insertion,
 * so no split and no `<w:del>` — though every non-text child of the cut (a
 * note or comment reference, a drawing, a text-box anchor) is kept. A
 * `<w:moveTo>` is never merged into, even by its own author: moved text is
 * relocated original text, not the author's new words.
 *
 * Nested non-run children (another author's `<w:del>`, bookmarks) move WHOLE,
 * counted at their view-visible length, and are never wrapped in the new
 * `<w:del>`: one that falls inside the span keeps its place between the cut
 * runs, which are deleted as one `<w:del>` per contiguous group.
 */
export function replaceSpanInsideRevision(
	paragraph: XmlNode,
	wrapper: XmlNode,
	span: Span,
	replacement: string,
	runProperties: XmlNode | null,
	tracked: TrackedReplaceOptions,
	view: FindView,
	formatting?: ReplacementFormatting,
): void {
	const wrapperIndex = paragraph.children.indexOf(wrapper);
	const wrapperStart = sumVisibleTextLength(
		paragraph.children.slice(0, wrapperIndex),
		view,
	);
	const parts = partitionWrapperChildren(wrapper, wrapperStart, span, view);
	const runs = replacementRuns(runProperties, replacement, formatting);

	if (
		wrapper.tag === "w:ins" &&
		wrapper.attributes["w:author"] === tracked.meta.author
	) {
		wrapper.children = [
			...parts.pre,
			...runs,
			...cutRemnants(parts.cut),
			...parts.post,
		];
		return;
	}

	const markerTail = parts.post.every(isPureMarker);
	wrapper.children = [
		...parts.pre,
		...deleteCutInOrder(parts.cut, tracked),
		...(markerTail ? parts.post : []),
	];
	const editorInsertion = <Ins meta={mintMeta(tracked)}>{runs}</Ins>;
	const tail = markerTail ? [] : [postHalfOf(wrapper, parts.post, tracked)];
	paragraph.children.splice(wrapperIndex + 1, 0, editorInsertion, ...tail);
}

/** Split the wrapper's children around `span` using the same offset space
 *  `collectRunSlots` walked: runs are sliced into pre / cut / post; any other
 *  child moves whole, counted at its visible length — one ending at or before
 *  the span start goes to `pre`, one starting at or after the span end to
 *  `post`, and one inside the span stays in `cut` at its document position. */
function partitionWrapperChildren(
	wrapper: XmlNode,
	wrapperStart: number,
	span: Span,
	view: FindView,
): { pre: XmlNode[]; cut: XmlNode[]; post: XmlNode[] } {
	const pre: XmlNode[] = [];
	const cut: XmlNode[] = [];
	const post: XmlNode[] = [];
	let offset = wrapperStart;

	for (const child of wrapper.children) {
		const length =
			child.tag === "w:r"
				? runTextLength(child)
				: isWrapperVisibleInView(child.tag, view)
					? sumVisibleTextLength(wrapperContent(child), view)
					: 0;
		const childStart = offset;
		const childEnd = offset + length;
		offset = childEnd;

		if (childEnd <= span.start) {
			pre.push(child);
			continue;
		}
		if (childStart >= span.end) {
			post.push(child);
			continue;
		}
		if (child.tag !== "w:r") {
			cut.push(child);
			continue;
		}

		const sliceStartInRun = Math.max(0, span.start - childStart);
		const sliceEndInRun = Math.min(length, span.end - childStart);
		if (sliceStartInRun > 0) pre.push(sliceRun(child, 0, sliceStartInRun));
		cut.push(sliceRun(child, sliceStartInRun, sliceEndInRun));
		if (sliceEndInRun < length) {
			post.push(sliceRun(child, sliceEndInRun, length));
		}
	}

	return { pre, cut, post };
}

/** What the editor's own cut leaves behind, in order: every non-run child,
 *  and each cut run stripped of its text — a run that still carries a child
 *  beyond `<w:rPr>` (a note or comment reference, a drawing, a text-box
 *  anchor) stays, so whatever was anchored on the replaced words survives;
 *  a text-only run is dropped. */
function cutRemnants(cut: XmlNode[]): XmlNode[] {
	const out: XmlNode[] = [];
	for (const child of cut) {
		if (child.tag !== "w:r") {
			out.push(child);
			continue;
		}
		child.children = child.children.filter(
			(part) => part.tag !== "w:t" && part.tag !== "w:delText",
		);
		if (child.children.some((part) => part.tag !== "w:rPr")) out.push(child);
	}
	return out;
}

/** The cut, deleted in document order: each contiguous group of runs becomes
 *  one editor `<w:del>` (text → delText); a non-run child between groups
 *  passes through untouched, where it was. */
function deleteCutInOrder(
	cut: XmlNode[],
	tracked: TrackedReplaceOptions,
): XmlNode[] {
	const out: XmlNode[] = [];
	let group: XmlNode[] = [];
	const flush = (): void => {
		if (group.length === 0) return;
		out.push(<Del meta={mintMeta(tracked)}>{group}</Del>);
		group = [];
	};
	for (const child of cut) {
		if (child.tag === "w:r") {
			group.push(convertTextToDelText(child));
			continue;
		}
		flush();
		out.push(child);
	}
	flush();
	return out;
}

const PURE_MARKER_TAGS: ReadonlySet<string> = new Set([
	"w:bookmarkStart",
	"w:bookmarkEnd",
	"w:commentRangeStart",
	"w:commentRangeEnd",
	"w:permStart",
	"w:permEnd",
	"w:proofErr",
]);

/** A range/annotation marker that is not revision content: one of the
 *  paragraph-level marker tags, or a run holding only a comment reference. */
function isPureMarker(node: XmlNode): boolean {
	if (PURE_MARKER_TAGS.has(node.tag)) return true;
	if (node.tag !== "w:r") return false;
	return node.children.every(
		(part) => part.tag === "w:rPr" || part.tag === "w:commentReference",
	);
}

/** The tail of a split wrapper: same tag, author and date, fresh `w:id`
 *  (two revisions may not share one). */
function postHalfOf(
	wrapper: XmlNode,
	post: XmlNode[],
	tracked: TrackedReplaceOptions,
): XmlNode {
	const half = new XmlNode(wrapper.tag, {
		...wrapper.attributes,
		"w:id": String(tracked.allocator.next()),
	});
	half.children = post;
	return half;
}

function mintMeta(tracked: TrackedReplaceOptions): TrackedMeta {
	return { ...tracked.meta, revisionId: tracked.allocator.next() };
}

/** The gate `replaceSpanInParagraph` evaluates: whether `node` is a
 *  `<w:ins>` / `<w:moveTo>` sitting directly under `paragraph`. A wrapper
 *  nested in a hyperlink, smartTag or field stays on the general walkers. */
export function isDirectRevisionWrapper(
	paragraph: XmlNode,
	node: XmlNode,
): boolean {
	if (node.tag !== "w:ins" && node.tag !== "w:moveTo") return false;
	return paragraph.children.includes(node);
}
