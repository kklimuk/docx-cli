import {
	isWrapperVisibleInView,
	rewrapSplitHalf,
	runTextLength,
	sumVisibleTextLength,
	wrapperContent,
	XmlNode,
} from "../../parser";
import { Ins } from "../../track-changes/emit";
import { replacementRuns } from "./replacement-runs";
import {
	mintMeta,
	mintTailRevisionId,
	partitionAroundSpan,
	type SpanContext,
	settleCut,
	splitAroundSpan,
} from "./split";

/**
 * Rebuild `container`'s children around `context.span` (paragraph offsets;
 * `baseOffset` is where the container's content begins): runs are sliced at
 * the span's edges, every visible wrapper the span crosses is split around it
 * (`splitAroundSpan`, recursively), the cut is settled (`settleCut`) and the
 * replacement placed once, at the cut, at THIS level — never inside a wrapper
 * it crossed, where it would inherit another author's revision or the
 * wrapper's semantics. The one exception is a link the span STARTS inside:
 * the replacement joins its leading half (`splitHyperlinkAcrossSpan`).
 * Wrappers hidden in the view, and non-wrappers (pPr, bookmarks, …), pass
 * through untouched: their text adds nothing to the offset and the span never
 * slices into them.
 */
export function rebuildAroundSpan(
	container: XmlNode,
	baseOffset: number,
	context: SpanContext,
): void {
	const { span, view } = context;
	const out: XmlNode[] = [];
	let offset = baseOffset;
	let placed = false;
	const takeReplacement = (): XmlNode[] => {
		if (placed) return [];
		placed = true;
		return placedReplacement(context);
	};

	for (const child of container.children) {
		if (child.tag === "w:r") {
			const sides = partitionAroundSpan([child], offset, context);
			offset += runTextLength(child);
			out.push(...sides.pre, ...settleCut(sides.cut, context));
			// The replacement goes at the cut — or, for an insertion point, where
			// the point falls: inside this run, or right before a run starting
			// at or after it.
			if (offset > span.start) out.push(...takeReplacement());
			out.push(...sides.post);
			continue;
		}
		if (!isWrapperVisibleInView(child.tag, view)) {
			out.push(child);
			continue;
		}

		const wrapperStart = offset;
		offset += sumVisibleTextLength(wrapperContent(child), view);
		if (offset <= span.start) {
			out.push(child);
			continue;
		}
		if (wrapperStart >= span.end) {
			out.push(...takeReplacement(), child);
			continue;
		}
		if (child.tag === "w:hyperlink") {
			out.push(
				...splitHyperlinkAcrossSpan(
					child,
					wrapperStart,
					offset,
					context,
					takeReplacement,
				),
			);
			continue;
		}
		// A transparent wrapper's attributes (e.g. w:fldSimple's w:instr) ride
		// both halves — splitting a fldSimple duplicates its instruction, but
		// Word re-evaluates fields on render and anything else would silently
		// drop the user's replacement intent.
		const { head, tail } = splitAroundSpan(child, wrapperStart, context);
		out.push(...head, ...takeReplacement());
		// Minted after the replacement's <w:ins>, so ids follow document order.
		mintTailRevisionId(child, head, tail, context.tracked);
		out.push(...tail);
	}

	out.push(...takeReplacement());
	container.children = out;
}

/** The replacement as it lands: its runs, inside its own `<w:ins>` under
 *  tracking, inside a copy of the link the span sat in when the level rose
 *  above it. */
function placedReplacement({ tracked, replacement }: SpanContext): XmlNode[] {
	const runs = replacementRuns(
		replacement.runProperties,
		replacement.text,
		replacement.formatting,
	);
	const nodes = tracked ? [<Ins meta={mintMeta(tracked)}>{runs}</Ins>] : runs;
	if (!replacement.linkShell) return nodes;
	const link = new XmlNode("w:hyperlink", {
		...replacement.linkShell.attributes,
	});
	link.children = nodes;
	return [link];
}

/** A link the span crosses at the rebuild level splits into two links around
 *  the cut — the cut stays inside, as a `<w:del>` under tracking so reject
 *  restores it. A span that STARTS inside the link puts the replacement in
 *  its leading half, so the replacement keeps the link. */
function splitHyperlinkAcrossSpan(
	link: XmlNode,
	linkStart: number,
	linkEnd: number,
	context: SpanContext,
	takeReplacement: () => XmlNode[],
): XmlNode[] {
	const { span } = context;
	const startsInside = span.start > linkStart && span.start < linkEnd;
	const { pre, cut, post } = partitionAroundSpan(
		link.children,
		linkStart,
		context,
	);
	const head = [...pre, ...settleCut(cut, context)];
	if (startsInside) head.push(...takeReplacement());
	return [
		...rewrapSplitHalf(link, head),
		...(startsInside ? [] : takeReplacement()),
		...rewrapSplitHalf(link, post),
	];
}
