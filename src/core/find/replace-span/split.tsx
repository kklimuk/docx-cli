import {
	isSubtractiveTrackedChangeWrapper,
	isTrackedChangeWrapper,
	isWrapperVisibleInView,
	rewrapSplitHalf,
	runTextLength,
	sliceRun,
	sumVisibleTextLength,
	wrapperContent,
	type XmlNode,
} from "../../parser";
import { type TrackedMeta, wrapContiguousTrackable } from "../../track-changes";
import { Del } from "../../track-changes/emit";
import type { FindView } from "../index";
import type { ReplacementFormatting, Span, TrackedReplaceOptions } from ".";

/** Split `wrapper` where its content crosses the span: the head keeps what
 *  lies before the span plus the settled cut, the tail what lies after. Cut
 *  content inside an additive revision gets our `<w:del>` nested there,
 *  keeping that author's attribution; inside a deletion (here, or any
 *  wrapper above) it simply stays deleted. Each half carries the wrapper's
 *  tag and attributes — except an `<mc:AlternateContent>`, whose halves come
 *  back BARE (`rewrapSplitHalf`). The caller gives a revision's tail its own
 *  `w:id` (`mintTailRevisionId`). */
export function splitAroundSpan(
	wrapper: XmlNode,
	wrapperStart: number,
	context: SpanContext,
): { head: XmlNode[]; tail: XmlNode[] } {
	const inner: SpanContext = {
		...context,
		deleted: context.deleted || isSubtractiveTrackedChangeWrapper(wrapper.tag),
	};
	const { pre, cut, post } = partitionAroundSpan(
		wrapperContent(wrapper),
		wrapperStart,
		inner,
	);
	return {
		head: rewrapSplitHalf(wrapper, [...pre, ...settleCut(cut, inner)]),
		tail: rewrapSplitHalf(wrapper, post),
	};
}

/** Both halves of a split revision are live revisions, so under tracking the
 *  tail gets its own `w:id` (same author/date) — a duplicate would corrupt
 *  the revision list. */
export function mintTailRevisionId(
	wrapper: XmlNode,
	head: XmlNode[],
	tail: XmlNode[],
	tracked: TrackedReplaceOptions | null,
): void {
	const [tailWrapper] = tail;
	if (
		tracked &&
		head.length > 0 &&
		tailWrapper &&
		isTrackedChangeWrapper(wrapper.tag) &&
		"w:id" in wrapper.attributes
	) {
		tailWrapper.attributes["w:id"] = String(tracked.allocator.next());
	}
}

/** Partition `children` (content starting at paragraph offset `start`)
 *  around the span. Runs are sliced at the span's edges. A nested visible
 *  wrapper the span reaches into (a link or smart tag inside another author's
 *  insertion, an insertion Word nested in a link) is split the same way,
 *  recursively — moved whole, the matched text it holds would survive the
 *  replace. Any other child can't be sliced, so it moves whole to the side
 *  `sideFor` picks, while its visible text still advances the offset so the
 *  runs after it stay aligned with the AST's accounting. */
export function partitionAroundSpan(
	children: XmlNode[],
	start: number,
	context: SpanContext,
): SpanSides {
	const { span, view } = context;
	const sides: SpanSides = { pre: [], cut: [], post: [] };
	let offset = start;
	for (const child of children) {
		if (child.tag !== "w:r") {
			const childStart = offset;
			offset += sumVisibleTextLength([child], view);
			const reachesInside =
				offset > childStart && offset > span.start && childStart < span.end;
			if (reachesInside && isWrapperVisibleInView(child.tag, view)) {
				const { head, tail } = splitAroundSpan(child, childStart, context);
				mintTailRevisionId(child, head, tail, context.tracked);
				// The head holds this wrapper's pre-span content (if any) and its
				// already-settled cut, which the caller's settleCut passes through.
				(childStart <= span.start ? sides.pre : sides.cut).push(...head);
				sides.post.push(...tail);
				continue;
			}
			sideFor(span, childStart, offset, sides).push(child);
			continue;
		}
		const length = runTextLength(child);
		const runStart = offset;
		offset += length;

		if (offset <= span.start) {
			sides.pre.push(child);
			continue;
		}
		if (runStart >= span.end) {
			sides.post.push(child);
			continue;
		}

		const sliceStart = Math.max(0, span.start - runStart);
		const sliceEnd = Math.min(length, span.end - runStart);
		if (sliceStart > 0) sides.pre.push(sliceRun(child, 0, sliceStart));
		const cutRun = sliceRun(child, sliceStart, sliceEnd);
		// An empty span (an insertion point inside a run) slices to a bare
		// `<w:r>` that must not become a phantom, empty tracked deletion.
		if (cutRun.children.some((part) => part.tag !== "w:rPr")) {
			sides.cut.push(cutRun);
		}
		if (sliceEnd < length) sides.post.push(sliceRun(child, sliceEnd, length));
	}
	return sides;
}

/** Which side of the span an unsliceable child covering `[start, end)` moves
 *  to, keeping document order: wholly before or after → that side; otherwise
 *  it is a zero-width marker inside the span (a bookmark, an equation,
 *  another author's deletion hidden in this view) → the cut, between the
 *  halves. A visible wrapper the span reaches into never gets here —
 *  `partitionAroundSpan` splits it instead — so nothing straddles an edge. */
function sideFor(
	span: Span,
	start: number,
	end: number,
	sides: SpanSides,
): XmlNode[] {
	if (end <= span.start) return sides.pre;
	if (start >= span.end) return sides.post;
	return sides.cut;
}

/** The cut content as it stays in the document. Untracked, the runs are
 *  dropped and anything unsliceable — another author's nested revision, a
 *  bookmark, an equation — survives. Already-deleted content (under a
 *  deletion here or above) stays as is. Under tracking each stretch of cut
 *  runs gets its own `<w:del>`, while the unsliceable children pass through
 *  in place, never nested in ours. */
export function settleCut(cut: XmlNode[], context: SpanContext): XmlNode[] {
	const { tracked } = context;
	if (!tracked) return cut.filter((node) => node.tag !== "w:r");
	if (context.deleted) return cut;
	return wrapContiguousTrackable(
		cut,
		(cutRuns) => {
			for (const cutRun of cutRuns) convertRunTextToDelText(cutRun);
			return <Del meta={mintMeta(tracked)}>{cutRuns}</Del>;
		},
		(node) => node.tag === "w:r",
	);
}

export function mintMeta(tracked: TrackedReplaceOptions): TrackedMeta {
	return { ...tracked.meta, revisionId: tracked.allocator.next() };
}

function convertRunTextToDelText(run: XmlNode): void {
	for (const child of run.children) {
		if (child.tag === "w:t") child.tag = "w:delText";
	}
}

/** Everything a rebuild needs besides the container. `deleted` is true when
 *  the content sits under a deletion — its cut is already gone on accept. */
export type SpanContext = {
	span: Span;
	view: FindView;
	tracked: TrackedReplaceOptions | null;
	deleted: boolean;
	replacement: {
		text: string;
		runProperties: XmlNode | null;
		formatting?: ReplacementFormatting;
		linkShell?: XmlNode;
	};
};

type SpanSides = { pre: XmlNode[]; cut: XmlNode[]; post: XmlNode[] };
