import type { Body } from "../ast/document/body";
import { w } from "../jsx";
import {
	isRunBearingWrapper,
	runTextLength,
	sliceRun,
	XmlNode,
} from "../parser";
import type { FindView, ParagraphSpanMatch } from "./index";
import {
	isWrapperVisibleInView,
	replacementRuns,
	sumVisibleTextLength,
} from "./replace-span";

/**
 * Editor-style replace of a span that may cross paragraph boundaries: the
 * REPLACEMENT's own newlines define the resulting paragraph structure, exactly
 * as if the span were selected in Word and the replacement typed. A
 * single-line replacement merges the spanned paragraphs into one (the first
 * paragraph's properties govern, and the last paragraph's mark — along with
 * any inline `<w:sectPr>` riding it — dies with the merge, as it does in
 * Word); a replacement containing newlines keeps/creates breaks: first
 * segment after the head, middle segments as fresh paragraphs cloned from the
 * first paragraph's properties, last segment before the tail, which keeps the
 * LAST paragraph's own properties. Works equally within one paragraph — a
 * multi-line replacement splits it.
 *
 * UNTRACKED only: OOXML records a paragraph merge/split as a paragraph-mark
 * `<w:ins>`/`<w:del>`, which this path does not emit yet — the CLI refuses
 * cross-paragraph replaces under tracking rather than record them wrong.
 *
 * Whole middle paragraphs are removed wholesale (the `delete --at pN`
 * precedent); the cut regions of the first/last paragraph salvage their
 * non-content children (comment/bookmark markers, view-invisible tracked
 * wrappers) the same way `replaceSpanInParagraph` passes them through, so a
 * cut can't orphan a comment anchor that started outside the span.
 *
 * The caller re-reads after this (block ids shift when paragraphs are
 * merged/split) — same contract as every structural mutation.
 */
export function replaceAcrossParagraphs(
	body: Body,
	match: ParagraphSpanMatch,
	replacement: string,
	view: FindView = "accepted",
): void {
	const startRef = body.resolveBlock(match.startBlockId);
	const endRef = body.resolveBlock(match.endBlockId);
	// Same-container / attached / forward-order validation via the canonical
	// range resolver every range mutator goes through.
	const { parent, startIndex, endIndex } = body.resolveBlockRange(
		match.startBlockId,
		match.endBlockId,
	);
	const sameNode = startRef.node === endRef.node;
	// CRLF-tolerant: a --batch JSONL replacement authored on Windows arrives as
	// real "\r\n" (JSON.parse decodes it; only inline argv goes through
	// decodeInlineEscapes' \r\n→\n mapping). Splitting on the pair keeps the
	// stray "\r" out of the emitted <w:t> runs.
	const segments = replacement.split(/\r\n|\n/);

	// Split the boundary paragraphs at the span edges. Same-node order matters:
	// split at the END offset first so the start split runs over an untouched
	// [0, endOffset) prefix.
	let head: XmlNode[];
	let tail: XmlNode[];
	let cut: XmlNode[];
	if (sameNode) {
		const endSplit = splitChildrenAt(
			startRef.node.children,
			match.endOffset,
			view,
		);
		const startSplit = splitChildrenAt(
			endSplit.before,
			match.startOffset,
			view,
		);
		head = startSplit.before;
		tail = endSplit.after;
		cut = startSplit.after;
	} else {
		const startSplit = splitChildrenAt(
			startRef.node.children,
			match.startOffset,
			view,
		);
		const endSplit = splitChildrenAt(
			endRef.node.children,
			match.endOffset,
			view,
		);
		head = startSplit.before;
		tail = endSplit.after;
		cut = [
			...startSplit.after,
			...endSplit.before.filter((child) => child.tag !== "w:pPr"),
		];
	}

	// Mirrors `replaceSpanInParagraph`'s firstSlot inheritance: the first
	// visible run in the cut region decides the rPr, even when it has none.
	const inherited =
		firstVisibleRun(cut, view)?.findChild("w:rPr")?.clone() ?? null;
	const salvaged = salvageNonContent(cut, view);
	const segmentRuns = (segment: string): XmlNode[] =>
		segment.length === 0 ? [] : replacementRuns(inherited, segment);

	const firstSegment = segments[0] ?? "";
	const startChildren = [...head, ...segmentRuns(firstSegment), ...salvaged];

	if (segments.length === 1) {
		// Merge: one paragraph keeps head + replacement + tail; the other spanned
		// paragraphs (their marks included) are removed.
		startRef.node.children = [...startChildren, ...tail];
		parent.splice(startIndex + 1, endIndex - startIndex);
		return;
	}

	startRef.node.children = startChildren;
	const template = paragraphPropertiesTemplate(startRef.node);
	const middle = segments
		.slice(1, -1)
		.map((segment) =>
			buildParagraph(template?.clone() ?? null, segmentRuns(segment)),
		);
	const lastSegment = segments[segments.length - 1] ?? "";

	if (sameNode) {
		// A split: the tail moves to a NEW final paragraph cloned from this one.
		const finalParagraph = buildParagraph(template?.clone() ?? null, [
			...segmentRuns(lastSegment),
			...tail,
		]);
		parent.splice(startIndex + 1, 0, ...middle, finalParagraph);
		return;
	}

	// The last paragraph survives with its own properties; middles are replaced
	// by the replacement's middle segments.
	const endProperties = endRef.node.findChild("w:pPr");
	endRef.node.children = [
		...(endProperties ? [endProperties] : []),
		...segmentRuns(lastSegment),
		...tail,
	];
	parent.splice(startIndex + 1, endIndex - startIndex - 1, ...middle);
}

type ChildSplit = { before: XmlNode[]; after: XmlNode[] };

/** Split a run container's children at a paragraph-local text offset, in the
 *  same offset space as `collectRunSlots`/`findTextSpans` (runs via
 *  `runTextLength`, visible wrappers descended and split when the offset lands
 *  inside them). `<w:pPr>` always stays in `before` (it's the paragraph's own
 *  metadata); zero-width children (markers, view-invisible tracked wrappers)
 *  go by position. */
function splitChildrenAt(
	children: XmlNode[],
	offset: number,
	view: FindView,
): ChildSplit {
	const before: XmlNode[] = [];
	const after: XmlNode[] = [];
	let cursor = 0;
	for (const child of children) {
		if (child.tag === "w:pPr") {
			before.push(child);
			continue;
		}
		if (child.tag === "w:r") {
			const length = runTextLength(child);
			if (cursor + length <= offset) {
				before.push(child);
			} else if (cursor >= offset) {
				after.push(child);
			} else {
				before.push(sliceRun(child, 0, offset - cursor));
				after.push(sliceRun(child, offset - cursor, length));
			}
			cursor += length;
			continue;
		}
		if (
			isRunBearingWrapper(child.tag) &&
			isWrapperVisibleInView(child.tag, view)
		) {
			const length = sumVisibleTextLength(child.children, view);
			if (cursor + length <= offset) {
				before.push(child);
			} else if (cursor >= offset) {
				after.push(child);
			} else {
				const inner = splitChildrenAt(child.children, offset - cursor, view);
				if (inner.before.length > 0) {
					const preWrapper = new XmlNode(child.tag, { ...child.attributes });
					preWrapper.children = inner.before;
					before.push(preWrapper);
				}
				if (inner.after.length > 0) {
					const postWrapper = new XmlNode(child.tag, { ...child.attributes });
					postWrapper.children = inner.after;
					after.push(postWrapper);
				}
			}
			cursor += length;
			continue;
		}
		// Zero-width: comment/bookmark markers, view-invisible tracked wrappers.
		(cursor < offset ? before : after).push(child);
	}
	return { before, after };
}

/** The first run in `children` that's visible in the view, descending into
 *  visible run-bearing wrappers. */
function firstVisibleRun(children: XmlNode[], view: FindView): XmlNode | null {
	for (const child of children) {
		if (child.tag === "w:r") return child;
		if (
			isRunBearingWrapper(child.tag) &&
			isWrapperVisibleInView(child.tag, view)
		) {
			const nested = firstVisibleRun(child.children, view);
			if (nested) return nested;
		}
	}
	return null;
}

/** What the cut region keeps: comment/bookmark markers (so a range that
 *  started outside the span isn't left dangling) and view-invisible tracked
 *  wrappers (whose content the span never covered — `replaceSpanInParagraph`
 *  passes them through untouched, so this path must too). Visible wrappers'
 *  own runs are consumed by the replacement; only markers inside them
 *  survive. */
function salvageNonContent(children: XmlNode[], view: FindView): XmlNode[] {
	const out: XmlNode[] = [];
	for (const child of children) {
		if (child.tag === "w:r" || child.tag === "w:pPr") continue;
		if (isRunBearingWrapper(child.tag)) {
			if (isWrapperVisibleInView(child.tag, view)) {
				out.push(...salvageNonContent(child.children, view));
			} else {
				out.push(child);
			}
			continue;
		}
		out.push(child);
	}
	return out;
}

/** The pPr for paragraphs minted by a split: the first paragraph's, minus the
 *  bits that must not duplicate — an inline `<w:sectPr>` (a second copy would
 *  mint a phantom section) and any paragraph-mark `<w:ins>`/`<w:del>` (a
 *  revision belongs to one mark, not every clone). */
function paragraphPropertiesTemplate(paragraph: XmlNode): XmlNode | null {
	const source = paragraph.findChild("w:pPr");
	if (!source) return null;
	const clone = source.clone();
	clone.children = clone.children.filter((child) => child.tag !== "w:sectPr");
	const markProperties = clone.findChild("w:rPr");
	if (markProperties) {
		markProperties.children = markProperties.children.filter(
			(child) => child.tag !== "w:ins" && child.tag !== "w:del",
		);
	}
	return clone;
}

function buildParagraph(
	paragraphProperties: XmlNode | null,
	runs: XmlNode[],
): XmlNode {
	return (
		<w.p>
			{paragraphProperties}
			{runs}
		</w.p>
	);
}
