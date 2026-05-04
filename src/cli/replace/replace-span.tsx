import { w } from "@core/jsx";
import {
	deepCloneNode,
	runTextLength,
	sliceRun,
	type XmlNode,
} from "@core/parser";

export type Span = { start: number; end: number };

export class TrackedChangeBoundaryError extends Error {
	constructor(message: string) {
		super(message);
		this.name = "TrackedChangeBoundaryError";
	}
}

/**
 * Replace text in a paragraph's runs at the given span with `replacement`.
 * Surrounding text and run formatting are preserved; the replacement run
 * inherits the rPr of the first run that overlaps the span.
 *
 * The span uses paragraph-relative offsets matching the AST's accounting,
 * which includes runs nested inside <w:ins>/<w:del>. When the span lies
 * fully inside a tracked-change wrapper, the replacement run is inserted
 * inside that wrapper (preserving the original change attribution). A span
 * that crosses a tracked-change boundary is rejected with
 * TrackedChangeBoundaryError — the agent should accept/reject the tracked
 * change first.
 */
export function replaceSpanInParagraph(
	paragraph: XmlNode,
	span: Span,
	replacement: string,
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
		paragraph.children.push(replacementRun(null, replacement));
		return;
	}

	const firstParent = firstSlot.parent;
	if (overlapping.some((slot) => slot.parent !== firstParent)) {
		throw new TrackedChangeBoundaryError(
			"Span crosses a tracked-change (<w:ins>/<w:del>) boundary; accept or reject the tracked change first.",
		);
	}

	const firstRunProperties = firstSlot.run.findChild("w:rPr");
	const inheritedProperties = firstRunProperties
		? deepCloneNode(firstRunProperties)
		: null;

	const containerStart = firstParent === paragraph ? 0 : firstSlot.offsetBefore;
	rebuildContainer(
		firstParent,
		containerStart,
		span,
		replacement,
		inheritedProperties,
		firstParent === paragraph,
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
		if (child.tag === "w:ins" || child.tag === "w:del") {
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
): void {
	const newChildren: XmlNode[] = [];
	let offset = baseOffset;
	let placedReplacement = false;

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
				if (!placedReplacement) {
					newChildren.push(replacementRun(runProperties, replacement));
					placedReplacement = true;
				}
				newChildren.push(child);
				continue;
			}

			const sliceStartInRun = Math.max(0, span.start - runStart);
			const sliceEndInRun = Math.min(length, span.end - runStart);
			if (sliceStartInRun > 0) {
				newChildren.push(sliceRun(child, 0, sliceStartInRun));
			}
			if (!placedReplacement) {
				newChildren.push(replacementRun(runProperties, replacement));
				placedReplacement = true;
			}
			if (sliceEndInRun < length) {
				newChildren.push(sliceRun(child, sliceEndInRun, length));
			}
			continue;
		}
		if (isParagraph && (child.tag === "w:ins" || child.tag === "w:del")) {
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

	if (!placedReplacement) {
		newChildren.push(replacementRun(runProperties, replacement));
	}

	container.children = newChildren;
}

function replacementRun(runProperties: XmlNode | null, text: string): XmlNode {
	if (text.length === 0) {
		// Preserve an empty run so downstream save logic doesn't trip on a
		// paragraph with zero runs; Word treats empty <w:r/> as harmless.
		return (
			<w.r>
				{runProperties}
				<w.t {...{ "xml:space": "preserve" }}>{""}</w.t>
			</w.r>
		);
	}
	return (
		<w.r>
			{runProperties}
			<w.t {...{ "xml:space": "preserve" }}>{text}</w.t>
		</w.r>
	);
}
