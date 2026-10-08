import { isTrackedChangeWrapper, type XmlNode } from "../../parser";
import type { Span, TrackedReplaceOptions } from ".";
import type { PathEntry, SpanTarget } from "./target";

/** The level the span is rebuilt at (see `replaceSpanInParagraph`). */
export function chooseLevel(
	target: SpanTarget,
	span: Span,
	tracked: TrackedReplaceOptions | undefined,
): Level {
	if (tracked) {
		const own = ownInsertionLevel(target, tracked.meta.author);
		if (own) return own;
	}
	// A tracked edit — and any pure insertion, which belongs beside a revision
	// rather than in it (an untracked one would join the other author's
	// revision and vanish when that is rejected) — rises above every revision.
	const rises = tracked !== undefined || span.start === span.end;
	const firstRevision = target.path.findIndex((entry) =>
		isTrackedChangeWrapper(entry.wrapper.tag),
	);
	if (!rises || firstRevision < 0) return { path: target.path };
	return {
		path: target.path.slice(0, firstRevision),
		linkShell: target.path
			.slice(firstRevision)
			.find((entry) => entry.wrapper.tag === "w:hyperlink")?.wrapper,
	};
}

/** The author's own insertion the span merges into, if any: the NEAREST
 *  revision on the span's path, with no other revision below it around any
 *  run the span touches. An insertion point at the edge of the own insertion
 *  merges too (typing at its end extends it), at the insertion's own level. */
function ownInsertionLevel(
	target: SpanTarget,
	author: string,
): Level | undefined {
	const nearest = nearestRevisionIndex(target.path);
	const insertion = target.path[nearest]?.wrapper;
	if (
		insertion &&
		isOwnInsertion(insertion, author) &&
		target.overlapping.every(
			(slot) => nearestRevisionIndex(slot.path) === nearest,
		)
	) {
		return { path: target.path, ownInsertion: insertion };
	}
	for (const neighbor of target.neighbors) {
		const index = nearestRevisionIndex(neighbor.path);
		const wrapper = neighbor.path[index]?.wrapper;
		if (wrapper && isOwnInsertion(wrapper, author)) {
			return {
				path: neighbor.path.slice(0, index + 1),
				ownInsertion: wrapper,
			};
		}
	}
	return undefined;
}

function nearestRevisionIndex(path: PathEntry[]): number {
	return path.findLastIndex((entry) =>
		isTrackedChangeWrapper(entry.wrapper.tag),
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

export type Level = {
	/** The wrappers above the rebuilt container, outermost first. */
	path: PathEntry[];
	/** Set when the span merges, untracked, into the author's own insertion. */
	ownInsertion?: XmlNode;
	/** A link the span sat in below the level, re-created around the
	 *  replacement so it keeps the link. */
	linkShell?: XmlNode;
};
