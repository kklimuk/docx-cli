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
 *  spans don't slice into them). Walkers recurse only into visible wrappers,
 *  so nested content counts only when EVERY enclosing revision is visible —
 *  the rule `isRevisionVisible` (core/ast/revision-visibility.ts) applies to
 *  the AST, which keeps find/replace in sync. (Exported for
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
 * tracked insertion in), which does exactly that. The one exception to the
 * exception: a span wholly inside the tracked author's OWN insertion (the
 * nearest enclosing `<w:ins>`/`<w:moveTo>` whose `w:author` matches exactly)
 * is rebuilt untracked inside it — the cut vanishes and the replacement joins
 * that insertion (same `w:id`/date), as Word does when you edit your own
 * pending text. A span that crosses out of it keeps the split/tracked shape.
 * A pure insertion at a revision wrapper's edge lands beside it, tracked or
 * not (unless it is the author's own insertion).
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
	// A tracked span inside a revision wrapper splits it (below). So does a pure
	// insertion at a wrapper's edge (`boundaryAnchor`), tracked or not: it
	// belongs beside that revision, not in it — an untracked one would join the
	// other author's revision and vanish when that revision is rejected.
	const splitsRevision =
		(tracked !== undefined || overlapping.length === 0) &&
		isTrackedChangeWrapper(firstParent.tag);
	// Editing your own pending insertion isn't a new revision: Word removes the
	// cut outright and types the replacement into that insertion — inside a
	// link nested in it too.
	const insideOwnInsertion =
		tracked !== undefined &&
		firstSlot.revision !== undefined &&
		isOwnInsertion(firstSlot.revision, tracked.meta.author);

	if (allSameParent && (!splitsRevision || insideOwnInsertion)) {
		rebuildContainer(
			firstParent,
			firstSlot.parentStart,
			span,
			replacement,
			inheritedProperties,
			insideOwnInsertion ? null : (tracked ?? null),
			view,
			formatting,
		);
		if (insideOwnInsertion && firstParent === firstSlot.revision) {
			dropIfBlank(firstParent, firstSlot.container);
		}
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

function collectRunSlots(paragraph: XmlNode, view: FindView): RunSlot[] {
	const slots: RunSlot[] = [];
	let offset = 0;
	function walk(
		parent: XmlNode,
		container: XmlNode,
		containerStart: number,
		revision: XmlNode | undefined,
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
					revision,
					run: child,
					offsetBefore: offset,
					length,
				});
				offset += length;
				continue;
			}
			if (isWrapperVisibleInView(child.tag, view)) {
				walk(
					wrapperContentNode(child),
					parent,
					parentStart,
					isTrackedChangeWrapper(child.tag) ? child : revision,
				);
			}
		}
	}
	walk(paragraph, paragraph, 0, undefined);
	return slots;
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
	/** The nearest enclosing revision wrapper (`<w:ins>`/`<w:del>`/
	 *  `<w:moveFrom>`/`<w:moveTo>`), if any — whose author decides whether a
	 *  tracked edit here is the author's own. */
	revision: XmlNode | undefined;
	run: XmlNode;
	offsetBefore: number;
	length: number;
};

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

/** An additive revision wrapper (`<w:ins>`/`<w:moveTo>`) written by `author`
 *  — exact string match on `w:author`, as Word compares reviewer names. */
function isOwnInsertion(wrapper: XmlNode, author: string): boolean {
	return (
		(wrapper.tag === "w:ins" || wrapper.tag === "w:moveTo") &&
		wrapper.getAttribute("w:author") === author
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

	// A revision wrapper is rebuilt here only untracked (the author's own
	// insertion) — under tracking those are split instead — so a tracked
	// replacement always gets its own <w:ins>, inside a hyperlink too, where it
	// used to land bare and untracked.
	const placeReplacement = (): void => {
		if (placed) return;
		placed = true;
		newChildren.push(
			...placedReplacement(runProperties, replacement, tracked, formatting),
		);
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

/** The replacement as it lands: its runs, inside its own `<w:ins>` under
 *  tracking. Shared by every path that places it. */
function placedReplacement(
	runProperties: XmlNode | null,
	replacement: string,
	tracked: TrackedReplaceOptions | null,
	formatting?: ReplacementFormatting,
): XmlNode[] {
	const runs = replacementRuns(runProperties, replacement, formatting);
	return tracked ? [<Ins meta={mintMeta(tracked)}>{runs}</Ins>] : runs;
}

/** Remove an own insertion the merge left holding nothing but blank runs —
 *  deleting all of your own pending text leaves no revision behind, as in
 *  Word, rather than a phantom empty `<w:ins>` that `track-changes list`
 *  would still report. */
function dropIfBlank(wrapper: XmlNode, container: XmlNode): void {
	if (!wrapper.children.every(isBlankRun)) return;
	const index = container.children.indexOf(wrapper);
	if (index >= 0) container.children.splice(index, 1);
}

function isBlankRun(node: XmlNode): boolean {
	return (
		node.tag === "w:r" &&
		node.children.every(
			(child) =>
				child.tag === "w:rPr" ||
				(child.tag === "w:t" && child.collectText().length === 0),
		)
	);
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
		newChildren.push(
			...placedReplacement(runProperties, replacement, tracked, formatting),
		);
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

		// Revision wrappers and transparent ones (w:fldSimple, w:smartTag,
		// mc:AlternateContent) split alike, with the replacement placed between
		// the halves at THIS level — never inside, where it would inherit
		// another author's revision or the wrapper's semantics. A transparent
		// wrapper's attributes (e.g. w:fldSimple's w:instr) ride both halves —
		// splitting a fldSimple duplicates its instruction, but Word re-evaluates
		// fields on render and anything else would silently drop the user's
		// replacement intent.
		const { head, tail } = splitAroundSpan(
			child,
			wrapperStart,
			span,
			tracked,
			view,
		);
		newChildren.push(...head);
		placeReplacement();
		// Minted after the replacement's <w:ins>, so ids follow document order.
		mintTailRevisionId(child, head, tail, tracked);
		newChildren.push(...tail);
	}

	if (!placed) placeReplacement();
	container.children = newChildren;
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
		tracked,
		view,
	);
	// Under tracking the cut link text stays as a <w:del> inside the link, so
	// reject restores it (it used to be dropped untracked).
	const preInner = [...pre, ...settleCut(cut, tracked, false)];

	if (startsInside) {
		// Replacement inherits the link: append it inside the pre-half.
		preInner.push(
			...placedReplacement(runProperties, replacement, tracked, formatting),
		);
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

/** Split `wrapper` where its content crosses `span`: the head keeps what lies
 *  before the span plus the settled cut, the tail what lies after. A
 *  subtractive wrapper's (w:del, w:moveFrom) cut simply stays deleted; an
 *  additive one's (w:ins, w:moveTo) gets our <w:del> nested inside, keeping
 *  that author's attribution. Each half carries the wrapper's tag and
 *  attributes — except an `<mc:AlternateContent>`, whose halves come back BARE
 *  (`rewrapSplitHalf`: it can't be halved into two valid Choice/Fallback
 *  pairs). The caller gives the tail of a revision its own `w:id`
 *  (`mintTailRevisionId`). */
function splitAroundSpan(
	wrapper: XmlNode,
	wrapperStart: number,
	span: Span,
	tracked: TrackedReplaceOptions | null,
	view: FindView,
): { head: XmlNode[]; tail: XmlNode[] } {
	const { pre, cut, post } = partitionAroundSpan(
		wrapperContent(wrapper),
		wrapperStart,
		span,
		tracked,
		view,
	);
	const head = rewrapSplitHalf(wrapper, [
		...pre,
		...settleCut(cut, tracked, isSubtractiveTrackedChangeWrapper(wrapper.tag)),
	]);
	return { head, tail: rewrapSplitHalf(wrapper, post) };
}

/** Both halves of a split revision are live revisions, so under tracking the
 *  tail gets its own `w:id` (same author/date) — a duplicate would corrupt
 *  the revision list. */
function mintTailRevisionId(
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

/** Partition a split wrapper's children around `span`. Runs are sliced at the
 *  span's edges. A nested visible wrapper the span reaches into (a link or
 *  smart tag inside another author's insertion, an insertion Word nested in a
 *  link) is split the same way, recursively — moved whole, the matched text
 *  it holds would survive the replace. Any other child can't be sliced, so it
 *  moves whole to the side `sideFor` picks, while its visible text still
 *  advances the offset so the runs after it stay aligned with the AST's
 *  accounting. */
function partitionAroundSpan(
	children: XmlNode[],
	start: number,
	span: Span,
	tracked: TrackedReplaceOptions | null,
	view: FindView,
): SpanSides {
	const sides: SpanSides = { pre: [], cut: [], post: [] };
	let offset = start;
	for (const child of children) {
		if (child.tag !== "w:r") {
			const childStart = offset;
			offset += sumVisibleTextLength([child], view);
			const reachesInside =
				offset > childStart && offset > span.start && childStart < span.end;
			if (reachesInside && isWrapperVisibleInView(child.tag, view)) {
				const { head, tail } = splitAroundSpan(
					child,
					childStart,
					span,
					tracked,
					view,
				);
				mintTailRevisionId(child, head, tail, tracked);
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
		if (hasRunContent(cutRun)) sides.cut.push(cutRun);
		if (sliceEnd < length) sides.post.push(sliceRun(child, sliceEnd, length));
	}
	return sides;
}

/** Which side of `span` an unsliceable child covering `[start, end)` moves to,
 *  keeping document order: wholly before or after → that side; wholly inside
 *  (a zero-width marker — a bookmark, an equation, another author's deletion
 *  hidden in this view) → the cut, between the halves; straddling an edge →
 *  the side of the edge it crosses. A visible wrapper the span reaches into
 *  never gets here — `partitionAroundSpan` splits it instead. */
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

type SpanSides = { pre: XmlNode[]; cut: XmlNode[]; post: XmlNode[] };

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
