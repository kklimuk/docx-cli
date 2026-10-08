import {
	isWrapperVisibleInView,
	runTextLength,
	wrapperContentNode,
	type XmlNode,
} from "../../parser";
import type { FindView } from "../index";
import type { Span } from ".";

/** Where the span sits: the wrapper path every run it touches shares. A
 *  non-empty span is anchored on the first run it overlaps. An empty span (a
 *  pure insertion, `pN:S-S`) inside a run overlaps that run; at a run
 *  boundary it overlaps nothing, so its neighbors decide — the path they
 *  share (inside a link only when BOTH neighbors are in it, so an insertion at
 *  a link's edge never silently extends the link). */
export function spanTarget(
	paragraph: XmlNode,
	span: Span,
	view: FindView,
): SpanTarget {
	const slots = collectRunSlots(paragraph, view);
	const overlapping = slots.filter(
		(slot) =>
			slot.offsetBefore + slot.length > span.start &&
			slot.offsetBefore < span.end,
	);
	if (overlapping.length > 0) {
		return {
			path: commonPath(overlapping.map((slot) => slot.path)),
			overlapping,
			neighbors: [],
		};
	}
	const before = slots.findLast(
		(slot) => slot.length > 0 && slot.offsetBefore + slot.length === span.start,
	);
	const after = slots.find(
		(slot) => slot.length > 0 && slot.offsetBefore === span.start,
	);
	const neighbors = [before, after].filter(
		(slot): slot is RunSlot => slot !== undefined,
	);
	const path = before && after ? commonPath([before.path, after.path]) : [];
	return { path, overlapping, neighbors };
}

/** Every run in `paragraph` visible in `view`, with its paragraph offset and
 *  the chain of run-bearing wrappers it sits in (outermost first). */
function collectRunSlots(paragraph: XmlNode, view: FindView): RunSlot[] {
	const slots: RunSlot[] = [];
	let offset = 0;
	function walk(content: XmlNode, path: PathEntry[]): void {
		for (const child of content.children) {
			if (child.tag === "w:r") {
				const length = runTextLength(child);
				slots.push({ path, run: child, offsetBefore: offset, length });
				offset += length;
				continue;
			}
			if (isWrapperVisibleInView(child.tag, view)) {
				walk(wrapperContentNode(child), [
					...path,
					{ wrapper: child, start: offset },
				]);
			}
		}
	}
	walk(paragraph, []);
	return slots;
}

function commonPath(paths: PathEntry[][]): PathEntry[] {
	const [first = [], ...rest] = paths;
	let length = first.length;
	for (const path of rest) {
		let shared = 0;
		while (
			shared < length &&
			shared < path.length &&
			path[shared]?.wrapper === first[shared]?.wrapper
		) {
			shared++;
		}
		length = shared;
	}
	return first.slice(0, length);
}

/** The run whose formatting the replacement inherits: the first overlapped
 *  run, or for an insertion point the neighbor before (else after) it — but
 *  never one inside a link the replacement won't sit in, whose style would
 *  dress plain text as a link. */
export function anchorFor(
	target: SpanTarget,
	depth: number,
	linkShell: XmlNode | undefined,
): RunSlot | undefined {
	if (target.overlapping.length > 0) return target.overlapping[0];
	return target.neighbors.find((slot) =>
		slot.path
			.slice(depth)
			.every(
				(entry) =>
					entry.wrapper.tag !== "w:hyperlink" || entry.wrapper === linkShell,
			),
	);
}

/** A run-bearing wrapper on a run's path, with the paragraph offset where its
 *  content begins. */
export type PathEntry = { wrapper: XmlNode; start: number };

export type RunSlot = {
	/** The wrappers enclosing the run, outermost first. */
	path: PathEntry[];
	run: XmlNode;
	offsetBefore: number;
	length: number;
};

export type SpanTarget = {
	/** The wrapper path shared by every run the span touches. */
	path: PathEntry[];
	overlapping: RunSlot[];
	/** An insertion point's neighbors (before, then after); empty otherwise. */
	neighbors: RunSlot[];
};
