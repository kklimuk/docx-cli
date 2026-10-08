import type { XmlNode } from "@core/parser";

/** Revision groups: a DERIVED VIEW over the existing `tcN` inventory that bundles
 *  an adjacent del+ins (or ins+del) text replace into one `revN` handle.
 *
 *  Word emits a separate `<w:del>` and `<w:ins>` for one logical replace (e.g.
 *  "Net 90" → "Net 30"). Accepting them one at a time forces a re-list between
 *  each, because `tcN` ids renumber after every accept — the `contract-finalize`
 *  friction ("tedious and error-prone"). `accept --at revN` resolves both halves
 *  in one call instead. This never touches the reader or the `tcN` ids; it only
 *  maps `revN` ⇄ its member `tcN`s, computed the SAME way in `list` and `apply`
 *  (sorted by tcN index, then paired) so the two can't disagree. */

type Groupable = {
	id: string;
	kind: string;
	author: string;
	date: string;
	blockId?: string;
	/** The revision's wrapper element — nesting is read off the tree. */
	node: XmlNode;
};

export type RevisionGroups = {
	/** `revN` → its member `tcN` ids (always exactly the two halves of a replace). */
	membersOf: Map<string, string[]>;
	/** `tcN` id → the `revN` it belongs to (only for grouped pairs). */
	revOf: Map<string, string>;
};

/** Pair every adjacent del+ins / ins+del on the same paragraph into a `revN`.
 *  Sorts by tcN index first so the pairing is identical regardless of input order
 *  (list builds its records differently than apply reads the inventory). */
export function revisionGroups(changes: Groupable[]): RevisionGroups {
	const sorted = [...changes].sort((a, b) => tcIndex(a.id) - tcIndex(b.id));
	const membersOf = new Map<string, string[]>();
	const revOf = new Map<string, string>();
	let next = 0;
	let index = 0;
	while (index < sorted.length) {
		const a = sorted[index];
		const b = sorted[index + 1];
		if (a && b && isTextReplacePair(a, b)) {
			const rev = `rev${next++}`;
			membersOf.set(rev, [a.id, b.id]);
			revOf.set(a.id, rev);
			revOf.set(b.id, rev);
			index += 2;
			continue;
		}
		index += 1;
	}
	return { membersOf, revOf };
}

/** An adjacent del+ins / ins+del on the same paragraph, by the same author
 *  at the same time (one edit), neither nested in the other — the canonical
 *  text "replace" shape. Restricted to plain `ins`/`del` (not moveFrom/moveTo,
 *  table, or section-property revisions) so grouping stays conservative;
 *  same-paragraph ins+ins / del+del (a paragraph insert/delete's run +
 *  paragraph-mark markers) never pair. A deletion nested INSIDE an insertion
 *  (the span-replace shape `<w:ins A><w:del …>…`, whoever deleted) is an edit
 *  of that insertion, not the other half of a replace; nor is the earlier
 *  insertion's split-off tail that follows it (same author, older date). */
function isTextReplacePair(a: Groupable, b: Groupable): boolean {
	if (a.blockId === undefined || a.blockId !== b.blockId) return false;
	if (a.author !== b.author || a.date !== b.date) return false;
	const pair = `${a.kind}+${b.kind}`;
	if (pair !== "del+ins" && pair !== "ins+del") return false;
	return !contains(a.node, b.node) && !contains(b.node, a.node);
}

function contains(ancestor: XmlNode, node: XmlNode): boolean {
	return ancestor.children.some(
		(child) => child === node || contains(child, node),
	);
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
