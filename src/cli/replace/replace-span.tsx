import { Del, Ins, type RevisionAllocator, type TrackedMeta } from "@core";
import { w } from "@core/jsx";
import { runTextLength, sliceRun, XmlNode } from "@core/parser";

export type Span = { start: number; end: number };

export type TrackedReplaceOptions = {
	meta: Omit<TrackedMeta, "revisionId">;
	allocator: RevisionAllocator;
};

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
 * surrounding wrapper's attribution.
 */
export function replaceSpanInParagraph(
	paragraph: XmlNode,
	span: Span,
	replacement: string,
	tracked?: TrackedReplaceOptions,
): void {
	if (span.start > span.end) {
		throw new Error(
			`replaceSpanInParagraph: invalid span ${span.start}-${span.end}`,
		);
	}

	const slots = collectRunSlots(paragraph);
	const overlapping = slots.filter(
		(slot) =>
			slot.offsetBefore + slot.length > span.start &&
			slot.offsetBefore < span.end,
	);

	const firstSlot = overlapping[0];
	if (!firstSlot) {
		paragraph.children.push(
			<ReplacementRun runProperties={null} text={replacement} />,
		);
		return;
	}

	const inheritedProperties = firstSlot.run.findChild("w:rPr")?.clone() ?? null;
	const firstParent = firstSlot.parent;
	const allSameParent = overlapping.every(
		(slot) => slot.parent === firstParent,
	);

	if (allSameParent) {
		const containerStart =
			firstParent === paragraph ? 0 : firstSlot.offsetBefore;
		rebuildContainer(
			firstParent,
			containerStart,
			span,
			replacement,
			inheritedProperties,
			firstParent === paragraph,
			tracked ?? null,
		);
		return;
	}

	rebuildAcrossBoundaries(
		paragraph,
		span,
		replacement,
		inheritedProperties,
		tracked ?? null,
	);
}

type RunSlot = {
	parent: XmlNode;
	run: XmlNode;
	offsetBefore: number;
	length: number;
};

function collectRunSlots(paragraph: XmlNode): RunSlot[] {
	const slots: RunSlot[] = [];
	let offset = 0;
	for (const child of paragraph.children) {
		if (child.tag === "w:r") {
			const length = runTextLength(child);
			slots.push({
				parent: paragraph,
				run: child,
				offsetBefore: offset,
				length,
			});
			offset += length;
			continue;
		}
		if (
			child.tag === "w:ins" ||
			child.tag === "w:del" ||
			child.tag === "w:hyperlink"
		) {
			for (const inner of child.children) {
				if (inner.tag !== "w:r") continue;
				const length = runTextLength(inner);
				slots.push({
					parent: child,
					run: inner,
					offsetBefore: offset,
					length,
				});
				offset += length;
			}
		}
	}
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
): void {
	const newChildren: XmlNode[] = [];
	let offset = baseOffset;
	let placed = false;

	const placeReplacement = (): void => {
		if (placed) return;
		placed = true;
		const run = (
			<ReplacementRun runProperties={runProperties} text={replacement} />
		);
		newChildren.push(
			tracked && isParagraph ? <Ins meta={mintMeta(tracked)}>{run}</Ins> : run,
		);
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
		if (
			isParagraph &&
			(child.tag === "w:ins" ||
				child.tag === "w:del" ||
				child.tag === "w:hyperlink")
		) {
			let innerLength = 0;
			for (const inner of child.children) {
				if (inner.tag === "w:r") innerLength += runTextLength(inner);
			}
			offset += innerLength;
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
): void {
	const newChildren: XmlNode[] = [];
	let offset = 0;
	let placed = false;

	const placeReplacement = (): void => {
		if (placed) return;
		placed = true;
		const run = (
			<ReplacementRun runProperties={runProperties} text={replacement} />
		);
		newChildren.push(tracked ? <Ins meta={mintMeta(tracked)}>{run}</Ins> : run);
	};

	for (const child of paragraph.children) {
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

		if (child.tag === "w:ins" || child.tag === "w:del") {
			const innerLength = sumInnerRunLengths(child);
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
			const innerLength = sumInnerRunLengths(child);
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

			splitHyperlinkAcrossSpan(
				child,
				wrapperStart,
				span,
				runProperties,
				replacement,
				tracked,
				newChildren,
				placeReplacement,
				() => {
					placed = true;
				},
			);
			continue;
		}

		newChildren.push(child);
	}

	if (!placed) placeReplacement();
	paragraph.children = newChildren;
}

function splitWrapperAcrossSpan(
	wrapper: XmlNode,
	wrapperStart: number,
	span: Span,
	tracked: TrackedReplaceOptions | null,
	out: XmlNode[],
	placeReplacement: () => void,
): void {
	const isDel = wrapper.tag === "w:del";
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
	if (isDel) {
		// Already deleted; the cut portion stays inside the original <w:del>.
		preChildren.push(...cutInner);
	} else if (tracked && cutInner.length > 0) {
		// <w:ins> overlap with tracking on: nest <w:del> inside, preserving original
		// author's insertion attribution.
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
		const postWrapper = new XmlNode(wrapper.tag, { ...wrapper.attributes });
		postWrapper.children = postInner;
		out.push(postWrapper);
	}
}

function splitHyperlinkAcrossSpan(
	wrapper: XmlNode,
	wrapperStart: number,
	span: Span,
	runProperties: XmlNode | null,
	replacement: string,
	tracked: TrackedReplaceOptions | null,
	out: XmlNode[],
	placeReplacement: () => void,
	markReplacementPlaced: () => void,
): void {
	const wrapperEnd =
		wrapperStart +
		wrapper.children.reduce(
			(total, inner) =>
				inner.tag === "w:r" ? total + runTextLength(inner) : total,
			0,
		);
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

	if (startsInside) {
		// Replacement inherits the link: append it inside the pre-half.
		const innerReplacement = (
			<ReplacementRun runProperties={runProperties} text={replacement} />
		);
		preInner.push(
			tracked ? (
				<Ins meta={mintMeta(tracked)}>{innerReplacement}</Ins>
			) : (
				innerReplacement
			),
		);
		markReplacementPlaced();
	}

	if (preInner.length > 0) {
		const preWrapper = new XmlNode("w:hyperlink", { ...wrapper.attributes });
		preWrapper.children = preInner;
		out.push(preWrapper);
	}

	if (!startsInside) placeReplacement();

	if (postInner.length > 0) {
		const postWrapper = new XmlNode("w:hyperlink", { ...wrapper.attributes });
		postWrapper.children = postInner;
		out.push(postWrapper);
	}
}

function sumInnerRunLengths(wrapper: XmlNode): number {
	let total = 0;
	for (const inner of wrapper.children) {
		if (inner.tag === "w:r") total += runTextLength(inner);
	}
	return total;
}

function mintMeta(tracked: TrackedReplaceOptions): TrackedMeta {
	return { ...tracked.meta, revisionId: tracked.allocator.next() };
}

function convertRunTextToDelText(run: XmlNode): void {
	for (const child of run.children) {
		if (child.tag === "w:t") child.tag = "w:delText";
	}
}

function ReplacementRun({
	runProperties,
	text,
}: {
	runProperties: XmlNode | null;
	text: string;
}): XmlNode {
	return (
		<w.r>
			{runProperties}
			<w.t {...{ "xml:space": "preserve" }}>{text}</w.t>
		</w.r>
	);
}
