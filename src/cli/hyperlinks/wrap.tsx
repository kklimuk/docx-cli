import { w } from "@core/jsx";
import { runTextLength, sliceRun, type XmlNode } from "@core/parser";

export type Span = { start: number; end: number };

export class HyperlinkWrapError extends Error {
	constructor(message: string) {
		super(message);
		this.name = "HyperlinkWrapError";
	}
}

/**
 * Wrap the runs inside `paragraph` that cover [span.start, span.end) in a
 * single `<w:hyperlink r:id="...">`. Splits runs at the span edges so the
 * surrounding text (and its run formatting) is preserved.
 *
 * Refuses to operate when the span overlaps an existing `<w:hyperlink>` —
 * nesting hyperlinks is not allowed in OOXML — or crosses `<w:ins>`/`<w:del>`
 * wrappers, since the run-splitting logic here is paragraph-level only.
 */
export function wrapSpanInHyperlink(
	paragraph: XmlNode,
	span: Span,
	relationshipId: string,
): void {
	if (span.start >= span.end) {
		throw new HyperlinkWrapError(
			`Empty or inverted span ${span.start}-${span.end}`,
		);
	}

	const newChildren: XmlNode[] = [];
	const wrappedRuns: XmlNode[] = [];
	let offset = 0;
	let placed = false;

	const placeWrapper = (): void => {
		if (placed || wrappedRuns.length === 0) return;
		newChildren.push(hyperlinkWrapper(relationshipId, wrappedRuns));
		wrappedRuns.length = 0;
		placed = true;
	};

	for (const child of paragraph.children) {
		if (child.tag === "w:r") {
			const length = runTextLength(child);
			const runStart = offset;
			const runEnd = offset + length;
			offset = runEnd;

			if (runEnd <= span.start || runStart >= span.end) {
				placeWrapper();
				newChildren.push(child);
				continue;
			}

			const sliceStartInRun = Math.max(0, span.start - runStart);
			const sliceEndInRun = Math.min(length, span.end - runStart);
			if (sliceStartInRun > 0) {
				newChildren.push(sliceRun(child, 0, sliceStartInRun));
			}
			wrappedRuns.push(sliceRun(child, sliceStartInRun, sliceEndInRun));
			if (runEnd >= span.end) placeWrapper();
			if (sliceEndInRun < length) {
				newChildren.push(sliceRun(child, sliceEndInRun, length));
			}
			continue;
		}

		if (child.tag === "w:hyperlink") {
			const hyperlinkLength = sumInnerRunLengths(child);
			const hyperlinkStart = offset;
			const hyperlinkEnd = offset + hyperlinkLength;
			offset = hyperlinkEnd;

			if (hyperlinkEnd <= span.start || hyperlinkStart >= span.end) {
				placeWrapper();
				newChildren.push(child);
				continue;
			}

			throw new HyperlinkWrapError(
				`Span ${span.start}-${span.end} overlaps an existing hyperlink at ${hyperlinkStart}-${hyperlinkEnd}; nested hyperlinks are not allowed`,
			);
		}

		if (child.tag === "w:ins" || child.tag === "w:del") {
			const innerLength = sumInnerRunLengths(child);
			const wrapperStart = offset;
			const wrapperEnd = offset + innerLength;
			offset = wrapperEnd;

			if (wrapperEnd <= span.start || wrapperStart >= span.end) {
				placeWrapper();
				newChildren.push(child);
				continue;
			}

			throw new HyperlinkWrapError(
				`Span ${span.start}-${span.end} crosses a tracked-change wrapper at ${wrapperStart}-${wrapperEnd}; resolve or accept the change first`,
			);
		}

		newChildren.push(child);
	}

	placeWrapper();
	paragraph.children = newChildren;
}

function hyperlinkWrapper(relationshipId: string, runs: XmlNode[]): XmlNode {
	return <w.hyperlink {...{ "r:id": relationshipId }}>{runs}</w.hyperlink>;
}

function sumInnerRunLengths(wrapper: XmlNode): number {
	let total = 0;
	for (const inner of wrapper.children) {
		if (inner.tag === "w:r") total += runTextLength(inner);
	}
	return total;
}
