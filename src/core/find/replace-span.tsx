import { textToRunElements } from "../blocks";
import { applyRunFormatToRpr, type RunFormat } from "../edit/set-formatting";
import { w } from "../jsx";
import {
	isRunBearingWrapper,
	isSubtractiveTrackedChangeWrapper,
	isTrackedChangeWrapper,
	rewrapSplitHalf,
	runTextLength,
	sliceRun,
	wrapperContent,
	wrapperContentNode,
	XmlNode,
} from "../parser";
import {
	type RevisionAllocator,
	type TrackedMeta,
	wrapContiguousTrackable,
} from "../track-changes";
import { Del, Ins } from "../track-changes/emit";
import type { FindView } from "./index";

export type Span = { start: number; end: number };

export type TrackedReplaceOptions = {
	meta: Omit<TrackedMeta, "revisionId">;
	allocator: RevisionAllocator;
};

export type ReplacementFormatting = {
	clearTags?: Set<string>;
	format?: RunFormat;
};

/** Whether a run-bearing wrapper's contents should be treated as VISIBLE in
 *  the chosen view. Invisible wrappers pass through replace's offset
 *  arithmetic untouched (their inner text adds nothing to the offset and
 *  spans don't slice into them). Mirrors `isRunVisibleInView` in
 *  src/core/find/index.ts so find/replace stay in sync. (Exported for
 *  replace-across.tsx, which walks the same offset space over whole
 *  paragraphs.) */
export function isWrapperVisibleInView(tag: string, view: FindView): boolean {
	if (!isRunBearingWrapper(tag)) return false;
	if (view === "current") return true;
	if (view === "accepted") return tag !== "w:del" && tag !== "w:moveFrom";
	return tag !== "w:ins" && tag !== "w:moveTo";
}

export function sumVisibleTextLength(
	children: XmlNode[],
	view: FindView,
): number {
	let total = 0;
	for (const child of children) {
		if (child.tag === "w:r") {
			total += runTextLength(child);
			continue;
		}
		if (isWrapperVisibleInView(child.tag, view)) {
			total += sumVisibleTextLength(wrapperContent(child), view);
		}
	}
	return total;
}

/**
 * Replace text in a paragraph's runs at the given span with `replacement`.
 * Surrounding text and run formatting are preserved; the replacement run
 * inherits the rPr of the first run that overlaps the span (for an empty span
 * at a run boundary — a pure insertion, `pN:S-S` — the run it is anchored on).
 *
 * The span uses paragraph-relative offsets matching the AST's accounting,
 * which includes runs nested inside <w:ins>/<w:del>. Spans may cross
 * tracked-change wrapper boundaries; overlapping wrappers are split into
 * pre/post halves so attribution survives unaffected portions.
 *
 * When `tracked` is provided, the cut content is wrapped in <w:del> (with
 * <w:t> nodes converted to <w:delText>), and the replacement is wrapped in
 * its own <w:ins>. When the span falls inside a non-revision wrapper
 * (same-parent case, e.g. a hyperlink), both stay inside it so the
 * replacement inherits the link.
 *
 * A tracked span wholly inside an existing revision wrapper (<w:ins>,
 * <w:moveTo>, <w:del>, <w:moveFrom>) is the exception (issue #13): left bare
 * inside it, the replacement would be attributed to THAT revision's author —
 * or, inside a deletion, vanish on accept. Word instead nests the cut's
 * <w:del> in the surrounding revision and splits it around the replacement's
 * own <w:ins>, so that case takes the across-boundaries path over the
 * wrapper's container (the paragraph, or e.g. the hyperlink Word nests a
 * tracked insertion in), which does exactly that.
 */
export function replaceSpanInParagraph(
	paragraph: XmlNode,
	span: Span,
	replacement: string,
	tracked?: TrackedReplaceOptions,
	view: FindView = "accepted",
	formatting?: ReplacementFormatting,
): void {
	if (span.start > span.end) {
		throw new Error(
			`replaceSpanInParagraph: invalid span ${span.start}-${span.end}`,
		);
	}

	const slots = collectRunSlots(paragraph, view);
	const overlapping = slots.filter(
		(slot) =>
			slot.offsetBefore + slot.length > span.start &&
			slot.offsetBefore < span.end,
	);

	const firstSlot = overlapping[0] ?? boundaryAnchor(slots, span);
	if (!firstSlot) {
		// Nothing to anchor on — an empty paragraph, or an insertion point at a
		// hyperlink's edge or inside one. Place it at paragraph level, at the
		// point, tracked like any other replacement.
		rebuildAcrossBoundaries(
			paragraph,
			0,
			span,
			replacement,
			null,
			tracked ?? null,
			view,
			formatting,
		);
		return;
	}

	const inheritedProperties = firstSlot.run.findChild("w:rPr")?.clone() ?? null;
	const firstParent = firstSlot.parent;
	const allSameParent = overlapping.every(
		(slot) => slot.parent === firstParent,
	);
	const insideTrackedRevision =
		tracked !== undefined && isTrackedChangeWrapper(firstParent.tag);

	if (allSameParent && !insideTrackedRevision) {
		rebuildContainer(
			firstParent,
			firstSlot.parentStart,
			span,
			replacement,
			inheritedProperties,
			tracked ?? null,
			view,
			formatting,
		);
		return;
	}

	// Across wrapper boundaries (from the paragraph down), or wholly inside one
	// revision wrapper under tracking — split within that wrapper's container.
	const container = allSameParent ? firstSlot.container : paragraph;
	rebuildAcrossBoundaries(
		container,
		allSameParent ? firstSlot.containerStart : 0,
		span,
		replacement,
		inheritedProperties,
		tracked ?? null,
		view,
		formatting,
	);
}

type RunSlot = {
	parent: XmlNode;
	/** Paragraph offset where `parent`'s content begins — NOT the offset of the
	 *  first overlapping run, which differs whenever the parent holds earlier
	 *  runs (or a nested deletion) ahead of the span. */
	parentStart: number;
	/** The node whose child list holds `parent` (the paragraph for a
	 *  paragraph-level wrapper, the hyperlink for an insertion Word nested in
	 *  one) and the offset where its content begins: where a revision wrapper
	 *  is split under tracking. */
	container: XmlNode;
	containerStart: number;
	run: XmlNode;
	offsetBefore: number;
	length: number;
};

function collectRunSlots(paragraph: XmlNode, view: FindView): RunSlot[] {
	const slots: RunSlot[] = [];
	let offset = 0;
	function walk(
		parent: XmlNode,
		container: XmlNode,
		containerStart: number,
	): void {
		const parentStart = offset;
		for (const child of parent.children) {
			if (child.tag === "w:r") {
				const length = runTextLength(child);
				slots.push({
					parent,
					parentStart,
					container,
					containerStart,
					run: child,
					offsetBefore: offset,
					length,
				});
				offset += length;
				continue;
			}
			if (isWrapperVisibleInView(child.tag, view)) {
				walk(wrapperContentNode(child), parent, parentStart);
			}
		}
	}
	walk(paragraph, paragraph, 0);
	return slots;
}

/** An empty span (a pure insertion, `pN:S-S`) at a run boundary overlaps no
 *  run. Anchor it on the run ending at the point — else the one starting
 *  there — so the text lands AT the point with that run's formatting (it used
 *  to be appended, untracked, at the paragraph's end). Only a run directly in
 *  the paragraph or in a paragraph-level revision wrapper anchors: a
 *  hyperlink's edge run would silently extend the link. */
function boundaryAnchor(slots: RunSlot[], span: Span): RunSlot | undefined {
	if (span.start !== span.end) return undefined;
	const anchorable = slots.filter(
		(slot) =>
			slot.length > 0 &&
			(slot.parent.tag === "w:p" ||
				(isTrackedChangeWrapper(slot.parent.tag) &&
					slot.container.tag === "w:p")),
	);
	return (
		anchorable.findLast(
			(slot) => slot.offsetBefore + slot.length === span.start,
		) ?? anchorable.find((slot) => slot.offsetBefore === span.start)
	);
}

function rebuildContainer(
	container: XmlNode,
	baseOffset: number,
	span: Span,
	replacement: string,
	runProperties: XmlNode | null,
	tracked: TrackedReplaceOptions | null,
	view: FindView,
	formatting?: ReplacementFormatting,
): void {
	const newChildren: XmlNode[] = [];
	let offset = baseOffset;
	let placed = false;

	// Under tracking the container is never a revision wrapper (those are
	// split instead), so the replacement always gets its own <w:ins> — inside
	// a hyperlink too, where it used to land bare and untracked.
	const placeReplacement = (): void => {
		if (placed) return;
		placed = true;
		const runs = replacementRuns(runProperties, replacement, formatting);
		if (tracked) {
			newChildren.push(<Ins meta={mintMeta(tracked)}>{runs}</Ins>);
			return;
		}
		newChildren.push(...runs);
	};

	for (const child of container.children) {
		if (child.tag === "w:r") {
			const runStart = offset;
			offset += runTextLength(child);
			pushRunAroundSpan(
				child,
				runStart,
				span,
				tracked,
				newChildren,
				placeReplacement,
			);
			continue;
		}
		// A nested visible wrapper (in a paragraph or a container alike) holds
		// no overlapping run — all of those are direct children here — but its
		// text still advances the offset, and an insertion point right before
		// it lands before it.
		if (isWrapperVisibleInView(child.tag, view)) {
			const wrapperStart = offset;
			offset += sumVisibleTextLength(wrapperContent(child), view);
			if (wrapperStart >= span.end) placeReplacement();
			newChildren.push(child);
			continue;
		}
		newChildren.push(child);
	}

	if (!placed) placeReplacement();
	container.children = newChildren;
}

/** Emit `run` into `out`: whole when it lies outside `span`, else sliced at
 *  the span's edges with the cut settled by `settleCut` (dropped, or a
 *  `<w:del>` under tracking) and the replacement placed at the cut. Shared by
 *  both rebuild paths. */
function pushRunAroundSpan(
	run: XmlNode,
	runStart: number,
	span: Span,
	tracked: TrackedReplaceOptions | null,
	out: XmlNode[],
	placeReplacement: () => void,
): void {
	const length = runTextLength(run);
	if (runStart + length <= span.start) {
		out.push(run);
		return;
	}
	if (runStart >= span.end) {
		placeReplacement();
		out.push(run);
		return;
	}
	const sliceStart = Math.max(0, span.start - runStart);
	const sliceEnd = Math.min(length, span.end - runStart);
	if (sliceStart > 0) out.push(sliceRun(run, 0, sliceStart));
	const cutRun = sliceRun(run, sliceStart, sliceEnd);
	if (hasRunContent(cutRun)) out.push(...settleCut([cutRun], tracked, false));
	placeReplacement();
	if (sliceEnd < length) out.push(sliceRun(run, sliceEnd, length));
}

/** A sliced run carries something besides its properties — an empty span
 *  (an insertion point inside a run) slices to a bare `<w:r>` that must not
 *  become a phantom, empty tracked deletion. */
function hasRunContent(run: XmlNode): boolean {
	return run.children.some((child) => child.tag !== "w:rPr");
}

/** The cut content as it stays in the document. Already-deleted content (cut
 *  from a subtractive wrapper) stays as is. Under tracking each stretch of cut
 *  runs gets its own <w:del>, while unsliceable children — another author's
 *  nested revision, a bookmark, an equation — pass through in place, never
 *  nested in ours. Untracked, the runs are dropped and those children
 *  survive. */
function settleCut(
	cut: XmlNode[],
	tracked: TrackedReplaceOptions | null,
	alreadyDeleted: boolean,
): XmlNode[] {
	if (alreadyDeleted) return cut;
	if (!tracked) return cut.filter((node) => node.tag !== "w:r");
	return wrapContiguousTrackable(
		cut,
		(cutRuns) => {
			for (const cutRun of cutRuns) convertRunTextToDelText(cutRun);
			return <Del meta={mintMeta(tracked)}>{cutRuns}</Del>;
		},
		(node) => node.tag === "w:r",
	);
}

function rebuildAcrossBoundaries(
	container: XmlNode,
	baseOffset: number,
	span: Span,
	replacement: string,
	runProperties: XmlNode | null,
	tracked: TrackedReplaceOptions | null,
	view: FindView,
	formatting?: ReplacementFormatting,
): void {
	const newChildren: XmlNode[] = [];
	let offset = baseOffset;
	let placed = false;

	const placeReplacement = (): void => {
		if (placed) return;
		placed = true;
		const runs = replacementRuns(runProperties, replacement, formatting);
		if (tracked) {
			newChildren.push(<Ins meta={mintMeta(tracked)}>{runs}</Ins>);
			return;
		}
		newChildren.push(...runs);
	};

	for (const child of container.children) {
		if (child.tag === "w:r") {
			const runStart = offset;
			offset += runTextLength(child);
			pushRunAroundSpan(
				child,
				runStart,
				span,
				tracked,
				newChildren,
				placeReplacement,
			);
			continue;
		}

		// Non-wrappers (pPr, bookmarks, …) and tracked-change wrappers invisible
		// in the chosen view pass through untouched — their inner text
		// contributes nothing to the offset and the span never slices into them.
		if (!isWrapperVisibleInView(child.tag, view)) {
			newChildren.push(child);
			continue;
		}

		const wrapperStart = offset;
		offset += sumVisibleTextLength(wrapperContent(child), view);
		if (offset <= span.start) {
			newChildren.push(child);
			continue;
		}
		if (wrapperStart >= span.end) {
			placeReplacement();
			newChildren.push(child);
			continue;
		}

		if (isTrackedChangeWrapper(child.tag)) {
			splitWrapperAcrossSpan(
				child,
				wrapperStart,
				span,
				tracked,
				view,
				newChildren,
				placeReplacement,
			);
			continue;
		}

		if (child.tag === "w:hyperlink") {
			splitHyperlinkAcrossSpan(
				child,
				wrapperStart,
				offset,
				span,
				runProperties,
				replacement,
				tracked,
				view,
				formatting,
				newChildren,
				placeReplacement,
				() => {
					placed = true;
				},
			);
			continue;
		}

		// Transparent wrappers (w:fldSimple, w:smartTag, mc:AlternateContent):
		// contents contribute to offset and may be split. Their attributes
		// (e.g. w:fldSimple's w:instr) are preserved on both halves of any
		// split — splitting a fldSimple would technically duplicate the field
		// instruction, but Word re-evaluates fields on next render and any
		// other behavior would silently drop the user's replacement intent.
		splitTransparentWrapperAcrossSpan(
			child,
			wrapperStart,
			span,
			tracked,
			view,
			newChildren,
			placeReplacement,
		);
	}

	if (!placed) placeReplacement();
	container.children = newChildren;
}

function splitWrapperAcrossSpan(
	wrapper: XmlNode,
	wrapperStart: number,
	span: Span,
	tracked: TrackedReplaceOptions | null,
	view: FindView,
	out: XmlNode[],
	placeReplacement: () => void,
): void {
	// Subtractive wrappers (w:del, w:moveFrom) hold content that's already
	// considered deleted — the cut portion stays in the pre-half wrapper.
	// Additive wrappers (w:ins, w:moveTo) hold "live" content; under tracking
	// the cut needs a new <w:del> wrapper nested inside, preserving the
	// surrounding author's insert/move-to attribution.
	const { pre, cut, post } = partitionAroundSpan(
		wrapper.children,
		wrapperStart,
		span,
		view,
	);
	const preChildren = [
		...pre,
		...settleCut(cut, tracked, isSubtractiveTrackedChangeWrapper(wrapper.tag)),
	];
	if (preChildren.length > 0) {
		const preWrapper = new XmlNode(wrapper.tag, { ...wrapper.attributes });
		preWrapper.children = preChildren;
		out.push(preWrapper);
	}

	placeReplacement();

	if (post.length > 0) {
		const postWrapper = new XmlNode(wrapper.tag, { ...wrapper.attributes });
		// Both halves are live revisions: the trailing one needs its own w:id
		// (same author/date) or the document carries a duplicate revision id.
		if (tracked && preChildren.length > 0 && "w:id" in wrapper.attributes) {
			postWrapper.attributes["w:id"] = String(tracked.allocator.next());
		}
		postWrapper.children = post;
		out.push(postWrapper);
	}
}

type SpanSides = { pre: XmlNode[]; cut: XmlNode[]; post: XmlNode[] };

/** Partition a split wrapper's children around `span`. Runs are sliced at the
 *  span's edges; any other child can't be sliced, so it moves whole to the
 *  side `sideFor` picks, while its visible text still advances the offset so
 *  the runs after it stay aligned with the AST's accounting. */
function partitionAroundSpan(
	children: XmlNode[],
	start: number,
	span: Span,
	view: FindView,
): SpanSides {
	const sides: SpanSides = { pre: [], cut: [], post: [] };
	let offset = start;
	for (const child of children) {
		if (child.tag !== "w:r") {
			const childStart = offset;
			offset += sumVisibleTextLength([child], view);
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
		if (hasRunContent(cutRun)) sides.cut.push(cutRun);
		if (sliceEnd < length) sides.post.push(sliceRun(child, sliceEnd, length));
	}
	return sides;
}

/** Which side of `span` an unsliceable child covering `[start, end)` moves to,
 *  keeping document order: wholly before or after → that side; wholly inside
 *  (a zero-width marker, or another author's nested deletion in the matched
 *  text) → the cut, between the halves; straddling an edge → the side of the
 *  edge it crosses. A child's own text is never re-cut, so visible text it
 *  holds inside the span survives the replace. */
function sideFor(
	span: Span,
	start: number,
	end: number,
	sides: SpanSides,
): XmlNode[] {
	if (end <= span.start) return sides.pre;
	if (start >= span.end) return sides.post;
	if (start >= span.start && end <= span.end) return sides.cut;
	return start < span.start ? sides.pre : sides.post;
}

/** Split a transparent wrapper (`<w:fldSimple>`, `<w:smartTag>`,
 * `<mc:AlternateContent>`) where its inner runs cross `span`. The cut is
 * settled like any other (dropped, or a `<w:del>` under tracking); pre/post
 * halves carry the wrapper's original attributes. The replacement run is
 * placed at top level (between pre and post halves) so it does not inherit
 * wrapper semantics. */
function splitTransparentWrapperAcrossSpan(
	wrapper: XmlNode,
	wrapperStart: number,
	span: Span,
	tracked: TrackedReplaceOptions | null,
	view: FindView,
	out: XmlNode[],
	placeReplacement: () => void,
): void {
	// `wrapperContent`: for an `<mc:AlternateContent>` this is its chosen
	// branch's runs — the halves come back BARE (`rewrapSplitHalf`), since a
	// wrapper can't be split into two valid Choice/Fallback pairs.
	const { pre, cut, post } = partitionAroundSpan(
		wrapperContent(wrapper),
		wrapperStart,
		span,
		view,
	);
	out.push(
		...rewrapSplitHalf(wrapper, [...pre, ...settleCut(cut, tracked, false)]),
	);
	placeReplacement();
	out.push(...rewrapSplitHalf(wrapper, post));
}

function splitHyperlinkAcrossSpan(
	wrapper: XmlNode,
	wrapperStart: number,
	wrapperEnd: number,
	span: Span,
	runProperties: XmlNode | null,
	replacement: string,
	tracked: TrackedReplaceOptions | null,
	view: FindView,
	formatting: ReplacementFormatting | undefined,
	out: XmlNode[],
	placeReplacement: () => void,
	markReplacementPlaced: () => void,
): void {
	const startsInside = span.start > wrapperStart && span.start < wrapperEnd;
	const { pre, cut, post } = partitionAroundSpan(
		wrapper.children,
		wrapperStart,
		span,
		view,
	);
	// Under tracking the cut link text stays as a <w:del> inside the link, so
	// reject restores it (it used to be dropped untracked).
	const preInner = [...pre, ...settleCut(cut, tracked, false)];

	if (startsInside) {
		// Replacement inherits the link: append it inside the pre-half.
		const innerRuns = replacementRuns(runProperties, replacement, formatting);
		if (tracked) {
			preInner.push(<Ins meta={mintMeta(tracked)}>{innerRuns}</Ins>);
		} else {
			preInner.push(...innerRuns);
		}
		markReplacementPlaced();
	}

	if (preInner.length > 0) {
		const preWrapper = new XmlNode("w:hyperlink", { ...wrapper.attributes });
		preWrapper.children = preInner;
		out.push(preWrapper);
	}

	if (!startsInside) placeReplacement();

	if (post.length > 0) {
		const postWrapper = new XmlNode("w:hyperlink", { ...wrapper.attributes });
		postWrapper.children = post;
		out.push(postWrapper);
	}
}

function mintMeta(tracked: TrackedReplaceOptions): TrackedMeta {
	return { ...tracked.meta, revisionId: tracked.allocator.next() };
}

function convertRunTextToDelText(run: XmlNode): void {
	for (const child of run.children) {
		if (child.tag === "w:t") child.tag = "w:delText";
	}
}

/** The replacement's run(s). Routed through `textToRunElements` so a real tab
 *  becomes `<w:tab/>` instead of a raw control character inside `<w:t>` — the
 *  same real-character handling every inline authoring surface uses. Plain
 *  single-line text still collapses to one `<w:t>` run, byte-identical to the
 *  old single-run shape; each produced run inherits the span's rPr. Returned
 *  as a plain array (never a fragment sentinel) so the nodes land directly in
 *  the live tree — `document.reread()` between batch entries must see them.
 *  (Real newlines never reach here — the CLI routes any `\n`-bearing
 *  replacement to the cross-paragraph path, where it means a paragraph mark.)
 *  Exported for replace-across.tsx, which builds each segment's runs the same
 *  way. */
export function replacementRuns(
	runProperties: XmlNode | null,
	text: string,
	formatting?: ReplacementFormatting,
): XmlNode[] {
	const properties = replacementRunProperties(runProperties, formatting);
	if (text.length === 0) {
		return [
			<w.r>
				{properties}
				<w.t {...{ "xml:space": "preserve" }} />
			</w.r>,
		];
	}
	const runs = textToRunElements(text);
	for (const run of runs) {
		if (properties) run.children.unshift(properties.clone());
	}
	return runs;
}

function replacementRunProperties(
	runProperties: XmlNode | null,
	formatting?: ReplacementFormatting,
): XmlNode | null {
	let properties = runProperties?.clone() ?? null;
	if (properties && formatting?.clearTags) {
		properties.children = properties.children.filter(
			(child) => !formatting.clearTags?.has(child.tag),
		);
		if (properties.children.length === 0) properties = null;
	}
	if (formatting?.format) {
		properties ??= XmlNode.element("w:rPr");
		applyRunFormatToRpr(properties, formatting.format);
	}
	return properties;
}
