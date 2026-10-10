import { textToRunElements } from "../blocks";
import { applyRunFormatToRpr, type RunFormat } from "../edit/set-formatting";
import { w } from "../jsx";
import {
	isRunBearingWrapper,
	isSubtractiveTrackedChangeWrapper,
	rewrapSplitHalf,
	runTextLength,
	sliceRun,
	wrapperContent,
	wrapperContentNode,
	XmlNode,
} from "../parser";
import type { RevisionAllocator, TrackedMeta } from "../track-changes";
import { Del, Ins } from "../track-changes/emit";
import type { FindView } from "./index";
import {
	isDirectRevisionWrapper,
	replaceSpanInsideRevision,
} from "./replace-in-revision";

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
 *  spans don't slice into them). The tag-side twin of `isRevisionVisible` in
 *  core/track-changes/revision-visibility.ts (the AST-side rule `find` reads);
 *  because a hidden wrapper is never descended, a wrapper nested inside a
 *  hidden one is hidden too — the same all-ancestors rule. (Exported for
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
 * inherits the rPr of the first run that overlaps the span.
 *
 * The span uses paragraph-relative offsets matching the AST's accounting,
 * which includes runs nested inside <w:ins>/<w:del>. Spans may cross
 * tracked-change wrapper boundaries; overlapping wrappers are split into
 * pre/post halves so attribution survives unaffected portions.
 *
 * When `tracked` is provided, the cut content is wrapped in <w:del> (with
 * <w:t> nodes converted to <w:delText>), and the replacement is wrapped in
 * <w:ins> at the paragraph top level. When the replacement falls inside an
 * existing wrapper (same-parent case), it stays unwrapped and inherits the
 * surrounding wrapper's attribution — except a tracked replace wholly inside
 * a paragraph-level <w:ins>/<w:moveTo>, which `replaceSpanInsideRevision`
 * splits (another author) or rebuilds in place (the editor's own).
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

	const firstSlot = overlapping[0];
	if (!firstSlot) {
		paragraph.children.push(...replacementRuns(null, replacement, formatting));
		return;
	}

	const inheritedProperties = firstSlot.run.findChild("w:rPr")?.clone() ?? null;
	const firstParent = firstSlot.parent;
	const allSameParent = overlapping.every(
		(slot) => slot.parent === firstParent,
	);

	// Tracked replace wholly inside another revision's <w:ins>/<w:moveTo>
	// (issue #13): Word's split / same-author merge, in replace-in-revision.
	if (
		tracked &&
		span.start < span.end &&
		allSameParent &&
		isDirectRevisionWrapper(paragraph, firstParent)
	) {
		replaceSpanInsideRevision(
			paragraph,
			firstParent,
			span,
			replacement,
			inheritedProperties,
			tracked,
			view,
			formatting,
		);
		return;
	}

	if (allSameParent) {
		rebuildContainer(
			firstParent,
			firstSlot.parentStart,
			span,
			replacement,
			inheritedProperties,
			firstParent === paragraph,
			tracked ?? null,
			view,
			formatting,
		);
		return;
	}

	rebuildAcrossBoundaries(
		paragraph,
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
	/** Offset at which `parent`'s content starts (0 for the paragraph). */
	parentStart: number;
	run: XmlNode;
	offsetBefore: number;
	length: number;
};

function collectRunSlots(paragraph: XmlNode, view: FindView): RunSlot[] {
	const slots: RunSlot[] = [];
	let offset = 0;
	function walk(parent: XmlNode, children: XmlNode[]): void {
		const parentStart = offset;
		for (const child of children) {
			if (child.tag === "w:r") {
				const length = runTextLength(child);
				slots.push({
					parent,
					parentStart,
					run: child,
					offsetBefore: offset,
					length,
				});
				offset += length;
				continue;
			}
			if (isWrapperVisibleInView(child.tag, view)) {
				const content = wrapperContentNode(child);
				walk(content, content.children);
			}
		}
	}
	walk(paragraph, paragraph.children);
	return slots;
}

function rebuildContainer(
	container: XmlNode,
	baseOffset: number,
	span: Span,
	replacement: string,
	runProperties: XmlNode | null,
	isParagraph: boolean,
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
		if (tracked && isParagraph) {
			newChildren.push(<Ins meta={mintMeta(tracked)}>{runs}</Ins>);
			return;
		}
		newChildren.push(...runs);
	};

	for (const child of container.children) {
		if (child.tag === "w:r") {
			const length = runTextLength(child);
			const runStart = offset;
			const runEnd = offset + length;
			offset = runEnd;

			if (runEnd <= span.start) {
				newChildren.push(child);
				continue;
			}
			if (runStart >= span.end) {
				placeReplacement();
				newChildren.push(child);
				continue;
			}

			const sliceStartInRun = Math.max(0, span.start - runStart);
			const sliceEndInRun = Math.min(length, span.end - runStart);
			if (sliceStartInRun > 0) {
				newChildren.push(sliceRun(child, 0, sliceStartInRun));
			}
			if (tracked) {
				const cutRun = sliceRun(child, sliceStartInRun, sliceEndInRun);
				convertRunTextToDelText(cutRun);
				newChildren.push(<Del meta={mintMeta(tracked)}>{cutRun}</Del>);
			}
			placeReplacement();
			if (sliceEndInRun < length) {
				newChildren.push(sliceRun(child, sliceEndInRun, length));
			}
			continue;
		}
		if (isWrapperVisibleInView(child.tag, view)) {
			offset += sumVisibleTextLength(wrapperContent(child), view);
			newChildren.push(child);
			continue;
		}
		newChildren.push(child);
	}

	if (!placed) placeReplacement();
	container.children = newChildren;
}

function rebuildAcrossBoundaries(
	paragraph: XmlNode,
	span: Span,
	replacement: string,
	runProperties: XmlNode | null,
	tracked: TrackedReplaceOptions | null,
	view: FindView,
	formatting?: ReplacementFormatting,
): void {
	const across: AcrossSpan = {
		span,
		replacement,
		runProperties,
		tracked,
		view,
		formatting,
		placed: false,
	};
	paragraph.children = rebuildChildrenAcrossSpan(
		paragraph.children,
		0,
		across,
		true,
	);
	if (!across.placed) {
		paragraph.children.push(...placedReplacement(across, true));
	}
}

/** One cross-boundary replace's shared state. `placed` is shared across the
 *  recursion so the replacement lands exactly once, in the container that
 *  holds the span's first character. */
type AcrossSpan = {
	span: Span;
	replacement: string;
	runProperties: XmlNode | null;
	tracked: TrackedReplaceOptions | null;
	view: FindView;
	formatting: ReplacementFormatting | undefined;
	placed: boolean;
};

/** The span removed from `children` (which start at offset `baseOffset`),
 *  with the replacement placed where the span starts. `isParagraph` gates the
 *  tracked replacement's own `<w:ins>`: only a paragraph-level replacement is
 *  wrapped (the recursion into a wrapper's content runs untracked only). */
function rebuildChildrenAcrossSpan(
	children: XmlNode[],
	baseOffset: number,
	across: AcrossSpan,
	isParagraph: boolean,
): XmlNode[] {
	const { span, tracked, view } = across;
	const newChildren: XmlNode[] = [];
	let offset = baseOffset;

	const placeReplacement = (): void => {
		if (across.placed) return;
		newChildren.push(...placedReplacement(across, isParagraph));
	};

	for (const child of children) {
		if (child.tag === "w:r") {
			const length = runTextLength(child);
			const runStart = offset;
			const runEnd = offset + length;
			offset = runEnd;

			if (runEnd <= span.start) {
				newChildren.push(child);
				continue;
			}
			if (runStart >= span.end) {
				placeReplacement();
				newChildren.push(child);
				continue;
			}

			const sliceStartInRun = Math.max(0, span.start - runStart);
			const sliceEndInRun = Math.min(length, span.end - runStart);
			if (sliceStartInRun > 0) {
				newChildren.push(sliceRun(child, 0, sliceStartInRun));
			}
			if (tracked) {
				const cutRun = sliceRun(child, sliceStartInRun, sliceEndInRun);
				convertRunTextToDelText(cutRun);
				newChildren.push(<Del meta={mintMeta(tracked)}>{cutRun}</Del>);
			}
			placeReplacement();
			if (sliceEndInRun < length) {
				newChildren.push(sliceRun(child, sliceEndInRun, length));
			}
			continue;
		}

		// Wrappers invisible in the chosen view (and any non-wrapper child)
		// pass through untouched — their inner text contributes nothing to the
		// offset and the span never slices into them.
		if (!isWrapperVisibleInView(child.tag, view)) {
			newChildren.push(child);
			continue;
		}

		const innerLength = sumVisibleTextLength(wrapperContent(child), view);
		const wrapperStart = offset;
		const wrapperEnd = offset + innerLength;
		offset = wrapperEnd;

		if (wrapperEnd <= span.start) {
			newChildren.push(child);
			continue;
		}
		if (wrapperStart >= span.end) {
			placeReplacement();
			newChildren.push(child);
			continue;
		}

		if (isTrackedChangeWrapper(child.tag)) {
			if (!tracked) {
				cutRevisionInPlace(
					child,
					wrapperStart,
					across,
					newChildren,
					placeReplacement,
				);
				continue;
			}
			splitWrapperAcrossSpan(
				child,
				wrapperStart,
				span,
				tracked,
				newChildren,
				placeReplacement,
			);
			continue;
		}

		if (child.tag === "w:hyperlink") {
			splitHyperlinkAcrossSpan(
				child,
				wrapperStart,
				across,
				newChildren,
				placeReplacement,
			);
			continue;
		}

		// Transparent wrappers (w:fldSimple, w:smartTag): contents contribute
		// to offset and may be split. Their attributes (e.g. w:fldSimple's
		// w:instr) are preserved on both halves of any split — splitting a
		// fldSimple would technically duplicate the field instruction, but
		// Word re-evaluates fields on next render and any other behavior
		// would silently drop the user's replacement intent.
		splitTransparentWrapperAcrossSpan(
			child,
			wrapperStart,
			span,
			newChildren,
			placeReplacement,
		);
	}

	return newChildren;
}

/** The replacement's runs, marking it placed. `wrapTracked` puts a tracked
 *  replacement in the editor's own `<w:ins>` (paragraph level, or inside a
 *  hyperlink it inherits); inside another revision's content it stays bare. */
function placedReplacement(
	across: AcrossSpan,
	wrapTracked: boolean,
): XmlNode[] {
	across.placed = true;
	const runs = replacementRuns(
		across.runProperties,
		across.replacement,
		across.formatting,
	);
	if (across.tracked && wrapTracked) {
		return [<Ins meta={mintMeta(across.tracked)}>{runs}</Ins>];
	}
	return runs;
}

function isTrackedChangeWrapper(tag: string): boolean {
	return (
		tag === "w:ins" ||
		tag === "w:del" ||
		tag === "w:moveFrom" ||
		tag === "w:moveTo"
	);
}

/** Untracked, a revision wrapper the span crosses is cut IN PLACE rather than
 *  split (issue #16): splitting it would put its `w:id` on both halves (an
 *  untracked replace has no revision-id allocator) and the replacement
 *  between them as plain original text. An additive wrapper (`<w:ins>` /
 *  `<w:moveTo>`) loses the span's text — descending nested hyperlinks,
 *  fields and revisions through the same walker — and keeps the replacement
 *  when the span starts inside it, credited to its author exactly as a match
 *  wholly inside it is; one left with no children is dropped. A subtractive
 *  wrapper's text is already deleted, so it stays whole and a replacement
 *  starting inside it lands right after it. */
function cutRevisionInPlace(
	wrapper: XmlNode,
	wrapperStart: number,
	across: AcrossSpan,
	out: XmlNode[],
	placeReplacement: () => void,
): void {
	if (isSubtractiveTrackedChangeWrapper(wrapper.tag)) {
		out.push(wrapper);
		placeReplacement();
		return;
	}
	wrapper.children = rebuildChildrenAcrossSpan(
		wrapper.children,
		wrapperStart,
		across,
		false,
	);
	if (wrapper.children.length > 0) out.push(wrapper);
}

function splitWrapperAcrossSpan(
	wrapper: XmlNode,
	wrapperStart: number,
	span: Span,
	tracked: TrackedReplaceOptions,
	out: XmlNode[],
	placeReplacement: () => void,
): void {
	// Subtractive wrappers (w:del, w:moveFrom) hold content that's already
	// considered deleted — the cut portion stays in the pre-half wrapper.
	// Additive wrappers (w:ins, w:moveTo) hold "live" content; the cut needs a
	// new <w:del> wrapper nested inside, preserving the surrounding author's
	// insert/move-to attribution. (Known gap: a wrapper NESTED in this one —
	// a hyperlink, another revision — rides the pre-half whole, uncounted.)
	const isSubtractive = isSubtractiveTrackedChangeWrapper(wrapper.tag);
	const preInner: XmlNode[] = [];
	const cutInner: XmlNode[] = [];
	const postInner: XmlNode[] = [];
	let innerOffset = wrapperStart;

	for (const inner of wrapper.children) {
		if (inner.tag !== "w:r") {
			preInner.push(inner);
			continue;
		}
		const length = runTextLength(inner);
		const runStart = innerOffset;
		const runEnd = innerOffset + length;
		innerOffset = runEnd;

		if (runEnd <= span.start) {
			preInner.push(inner);
			continue;
		}
		if (runStart >= span.end) {
			postInner.push(inner);
			continue;
		}

		const sliceStartInRun = Math.max(0, span.start - runStart);
		const sliceEndInRun = Math.min(length, span.end - runStart);
		if (sliceStartInRun > 0) preInner.push(sliceRun(inner, 0, sliceStartInRun));
		cutInner.push(sliceRun(inner, sliceStartInRun, sliceEndInRun));
		if (sliceEndInRun < length)
			postInner.push(sliceRun(inner, sliceEndInRun, length));
	}

	const preChildren = preInner.slice();
	if (isSubtractive) {
		preChildren.push(...cutInner);
	} else if (cutInner.length > 0) {
		for (const cutRun of cutInner) convertRunTextToDelText(cutRun);
		preChildren.push(<Del meta={mintMeta(tracked)}>{cutInner}</Del>);
	}
	if (preChildren.length > 0) {
		const preWrapper = new XmlNode(wrapper.tag, { ...wrapper.attributes });
		preWrapper.children = preChildren;
		out.push(preWrapper);
	}

	placeReplacement();

	if (postInner.length > 0) {
		// Two revisions can't share a w:id: when the pre-half kept it, the
		// post-half is a new revision with a fresh one.
		const postAttributes = { ...wrapper.attributes };
		if (preChildren.length > 0 && postAttributes["w:id"] !== undefined) {
			postAttributes["w:id"] = String(tracked.allocator.next());
		}
		const postWrapper = new XmlNode(wrapper.tag, postAttributes);
		postWrapper.children = postInner;
		out.push(postWrapper);
	}
}

/** Split a transparent wrapper (`<w:fldSimple>`, `<w:smartTag>`) where its
 * inner runs cross `span`. Cut content is dropped; pre/post halves carry the
 * wrapper's original attributes. The replacement run is placed at top level
 * (between pre and post halves) so it does not inherit wrapper semantics. */
function splitTransparentWrapperAcrossSpan(
	wrapper: XmlNode,
	wrapperStart: number,
	span: Span,
	out: XmlNode[],
	placeReplacement: () => void,
): void {
	const preInner: XmlNode[] = [];
	const postInner: XmlNode[] = [];
	let innerOffset = wrapperStart;

	// `wrapperContent`: for an `<mc:AlternateContent>` this is its chosen
	// branch's runs — the halves come back BARE (`rewrapSplitHalf`), since a
	// wrapper can't be split into two valid Choice/Fallback pairs.
	for (const inner of wrapperContent(wrapper)) {
		if (inner.tag !== "w:r") {
			preInner.push(inner);
			continue;
		}
		const length = runTextLength(inner);
		const runStart = innerOffset;
		const runEnd = innerOffset + length;
		innerOffset = runEnd;

		if (runEnd <= span.start) {
			preInner.push(inner);
			continue;
		}
		if (runStart >= span.end) {
			postInner.push(inner);
			continue;
		}

		const sliceStartInRun = Math.max(0, span.start - runStart);
		const sliceEndInRun = Math.min(length, span.end - runStart);
		if (sliceStartInRun > 0) preInner.push(sliceRun(inner, 0, sliceStartInRun));
		// cut content is dropped (replaced).
		if (sliceEndInRun < length)
			postInner.push(sliceRun(inner, sliceEndInRun, length));
	}

	out.push(...rewrapSplitHalf(wrapper, preInner));
	placeReplacement();
	out.push(...rewrapSplitHalf(wrapper, postInner));
}

function splitHyperlinkAcrossSpan(
	wrapper: XmlNode,
	wrapperStart: number,
	across: AcrossSpan,
	out: XmlNode[],
	placeReplacement: () => void,
): void {
	const { span, view } = across;
	const wrapperEnd =
		wrapperStart + sumVisibleTextLength(wrapper.children, view);
	const startsInside = span.start > wrapperStart && span.start < wrapperEnd;

	const preInner: XmlNode[] = [];
	const postInner: XmlNode[] = [];
	let innerOffset = wrapperStart;

	for (const inner of wrapper.children) {
		if (inner.tag !== "w:r") {
			preInner.push(inner);
			continue;
		}
		const length = runTextLength(inner);
		const runStart = innerOffset;
		const runEnd = innerOffset + length;
		innerOffset = runEnd;

		if (runEnd <= span.start) {
			preInner.push(inner);
			continue;
		}
		if (runStart >= span.end) {
			postInner.push(inner);
			continue;
		}

		const sliceStartInRun = Math.max(0, span.start - runStart);
		const sliceEndInRun = Math.min(length, span.end - runStart);
		if (sliceStartInRun > 0) preInner.push(sliceRun(inner, 0, sliceStartInRun));
		// cut portion is dropped (replaced by the replacement run)
		if (sliceEndInRun < length) {
			postInner.push(sliceRun(inner, sliceEndInRun, length));
		}
	}

	// Replacement inherits the link when the span starts inside it: append it
	// inside the pre-half (tracked, in the editor's own <w:ins>).
	if (startsInside && !across.placed) {
		preInner.push(...placedReplacement(across, true));
	}

	if (preInner.length > 0) {
		const preWrapper = new XmlNode("w:hyperlink", { ...wrapper.attributes });
		preWrapper.children = preInner;
		out.push(preWrapper);
	}

	placeReplacement();

	if (postInner.length > 0) {
		const postWrapper = new XmlNode("w:hyperlink", { ...wrapper.attributes });
		postWrapper.children = postInner;
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
