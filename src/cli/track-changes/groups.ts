/** Revision groups: a DERIVED VIEW over the existing `tcN` inventory that bundles
 *  one logical text replace into one `revN` handle.
 *
 *  Word emits a separate `<w:del>` and `<w:ins>` for one logical replace (e.g.
 *  "Net 90" → "Net 30"), and a replace whose matched text spans differently
 *  formatted runs emits SEVERAL `<w:del>`s (filling "[Fill in: date]" deletes "[",
 *  the underlined placeholder, and "]" as three wrappers around one insertion).
 *  Accepting them one at a time forces a re-list between each, because `tcN` ids
 *  renumber after every accept — the `contract-finalize` friction ("tedious and
 *  error-prone") — and rejecting only the first pair of a multi-run replace leaves
 *  the other deletions behind. `accept --at revN` resolves every member in one call
 *  instead. This never touches the reader or the `tcN` ids; it only maps `revN` ⇄
 *  its member `tcN`s, computed the SAME way in `list` and `apply` (sorted by tcN
 *  index, then chained) so the two can't disagree. */

import type { XmlNode } from "@core/parser";

type Groupable = {
	id: string;
	kind: string;
	blockId?: string;
	author: string;
	node: XmlNode;
	parent: XmlNode[];
};

export type RevisionGroups = {
	/** `revN` → its member `tcN` ids (every del/ins of one logical replace). */
	membersOf: Map<string, string[]>;
	/** `tcN` id → the `revN` it belongs to (only for grouped changes). */
	revOf: Map<string, string>;
};

/** Chain consecutive plain del/ins on the same paragraph by the same author into
 *  one `revN` when they sit side by side in the XML (or, for the classic two-part
 *  replace, are simply consecutive in tcN order), and keep the chain only if it
 *  holds both a deletion and an insertion — a run of pure deletions or pure
 *  insertions is not a replace. Sorts by tcN index first so the grouping is
 *  identical regardless of input order. */
export function revisionGroups(changes: Groupable[]): RevisionGroups {
	const sorted = [...changes].sort((a, b) => tcIndex(a.id) - tcIndex(b.id));
	const membersOf = new Map<string, string[]>();
	const revOf = new Map<string, string>();
	let next = 0;
	let index = 0;
	while (index < sorted.length) {
		const first = sorted[index];
		if (!first) break;
		const chain = [first];
		while (index + chain.length < sorted.length) {
			const previous = chain[chain.length - 1];
			const candidate = sorted[index + chain.length];
			if (
				!previous ||
				!candidate ||
				!continuesReplace(chain, previous, candidate)
			)
				break;
			chain.push(candidate);
		}
		const kinds = new Set(chain.map((change) => change.kind));
		if (chain.length >= 2 && kinds.has("del") && kinds.has("ins")) {
			const rev = `rev${next++}`;
			membersOf.set(
				rev,
				chain.map((change) => change.id),
			);
			for (const change of chain) revOf.set(change.id, rev);
			index += chain.length;
			continue;
		}
		index += 1;
	}
	return { membersOf, revOf };
}

/** Whether `candidate` extends the replace `chain` ends with `previous`. Only plain
 *  `ins`/`del` on the same paragraph by the same author qualify (not moveFrom/moveTo,
 *  paragraph-mark, table, or property revisions; one author's deletion nested in
 *  another's insertion is two decisions). The classic two-part pair groups on tcN
 *  adjacency alone, as it always has; anything longer must be XML-contiguous, so two
 *  separate edits in one paragraph never merge. A blockId that names a whole part
 *  rather than a paragraph (a note body's "", a header/footer's `hdrN`/`ftrN`)
 *  says nothing about "same paragraph", so there the two halves must also share a
 *  container — a deletion in one header line and an insertion in the next are two
 *  decisions, not one replace. */
function continuesReplace(
	chain: Groupable[],
	previous: Groupable,
	candidate: Groupable,
): boolean {
	if (!isPlainTextChange(previous) || !isPlainTextChange(candidate))
		return false;
	if (previous.blockId === undefined || previous.blockId !== candidate.blockId)
		return false;
	if (previous.author !== candidate.author) return false;
	if (isPartWide(previous.blockId) && previous.parent !== candidate.parent)
		return false;
	if (areSiblingsSideBySide(previous, candidate)) return true;
	return chain.length === 1 && previous.kind !== candidate.kind;
}

function isPlainTextChange(change: Groupable): boolean {
	return change.kind === "ins" || change.kind === "del";
}

function isPartWide(blockId: string): boolean {
	return blockId === "" || /^(hdr|ftr)\d+$/.test(blockId);
}

/** Markers that sit between runs without contributing text — a bookmark or comment
 *  range boundary between two halves of a replace doesn't split it. */
const TRANSPARENT_MARKERS = new Set([
	"w:proofErr",
	"w:bookmarkStart",
	"w:bookmarkEnd",
	"w:commentRangeStart",
	"w:commentRangeEnd",
	"w:permStart",
	"w:permEnd",
]);

function areSiblingsSideBySide(a: Groupable, b: Groupable): boolean {
	if (a.parent !== b.parent) return false;
	const from = a.parent.indexOf(a.node);
	const to = a.parent.indexOf(b.node);
	if (from < 0 || to <= from) return false;
	for (const between of a.parent.slice(from + 1, to)) {
		if (!TRANSPARENT_MARKERS.has(between.tag)) return false;
	}
	return true;
}

/** Expand any `revN` in a list of `--at` targets to its member `tcN` ids; pass
 *  `tcN` (and anything else) through untouched. Throws `UnknownRevisionError` on a
 *  `revN` that isn't a current group. */
export function expandRevisionTargets(
	targets: string[],
	groups: RevisionGroups,
): string[] {
	const out: string[] = [];
	for (const target of targets) {
		if (/^rev\d+$/.test(target)) {
			const members = groups.membersOf.get(target);
			if (!members) throw new UnknownRevisionError(target);
			out.push(...members);
			continue;
		}
		out.push(target);
	}
	return out;
}

export class UnknownRevisionError extends Error {
	constructor(public revision: string) {
		super(`Unknown revision group "${revision}"`);
		this.name = "UnknownRevisionError";
	}
}

function tcIndex(id: string): number {
	const match = id.match(/^tc(\d+)$/);
	return match?.[1] ? Number(match[1]) : 0;
}
