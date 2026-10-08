import type { SectionProperties, TrackedChange } from "@core";
import { Document, flattenParagraphs, readSectionProperties } from "@core";
import type { XmlNode } from "@core/parser";
import { TrackChanges } from "@core/track-changes";
import { revisionGroups } from "./groups";
import { renderTrackedChangeTable } from "./list-table";

/** Compact prior/current summary for a `pPrChange`: the direct paragraph
 *  properties (raw twips, matching `read --ast`). */
type ParagraphPropsSummary = {
	style?: string;
	alignment?: string;
	spacing?: Record<string, string | number>;
	indent?: Record<string, number>;
};

export type TrackedChangeRecord = Omit<TrackedChange, "within"> & {
	blockId: string;
	/** Every text run inside the wrapper, in document order — including the
	 *  text of revisions nested in it (an editor's deletion inside a reviewer's
	 *  insertion is still text that reviewer inserted). */
	text: string;
	prior?: SectionProperties | ParagraphPropsSummary;
	current?: SectionProperties | ParagraphPropsSummary;
	/** `revN` when this change is one half of a del+ins replace pair; absent for
	 *  solo changes. `accept/reject --at revN` acts on both halves at once. */
	group?: string;
	/** Table-only (stripped from `--json` by `listItem`): the ids of revisions
	 *  nested directly inside this one, and whether it is a paragraph-mark
	 *  marker — the only change whose empty text means "¶". */
	contains?: string[];
	paragraphMark?: boolean;
};

/** Walk a document's tracked changes into the enriched, `revN`-grouped records
 *  that both `track-changes list` and the post-accept/reject "remaining" view
 *  render. Sorted by tcN (document order). */
export function collectTrackedChangeRecords(
	document: Document,
): TrackedChangeRecord[] {
	const byId = new Map<string, TrackedChangeRecord>();
	const recordFor = (
		change: TrackedChange,
		blockId: string,
	): TrackedChangeRecord => {
		const existing = byId.get(change.id);
		if (existing) return existing;
		const { id, kind, author, date, revisionId } = change;
		const record = { id, kind, author, date, revisionId, blockId, text: "" };
		byId.set(id, record);
		return record;
	};
	for (const paragraph of flattenParagraphs(document.body.blocks)) {
		for (const run of paragraph.runs) {
			if (run.type !== "text") continue;
			// A run's text belongs to its innermost revision AND every revision
			// that revision nests in (the reader's `within` chain).
			let nested: TrackedChange | undefined;
			for (let change = run.trackedChange; change; change = change.within) {
				const record = recordFor(change, paragraph.id);
				record.text += run.text;
				if (nested && !record.contains?.includes(nested.id)) {
					record.contains = [...(record.contains ?? []), nested.id];
				}
				nested = change;
			}
		}
	}

	// The AST loop above only sees text-bearing run-level changes. Everything
	// else — paragraph-mark markers, section / table-property revisions,
	// checkbox toggles, and standalone note-body edits — comes from the single
	// tracked-change inventory the reader built (TrackChanges.list reads
	// document.trackedChangeReferences; no re-walk). kind, author, date and
	// revisionId are already resolved on each record.
	const inventory = new TrackChanges(document).list();
	for (const change of inventory) {
		if (byId.has(change.id)) continue;
		const record: TrackedChangeRecord = {
			id: change.id,
			kind: change.kind,
			author: change.author,
			date: change.date,
			revisionId: change.revisionId,
			blockId: change.blockId,
			text: "",
		};
		if (change.paragraph) record.paragraphMark = true;
		if (change.kind === "sectPrChange") {
			// Live siblings (parent array) carry the post-edit values; the
			// snapshot inside the change marker carries the prior values.
			const liveSiblings = change.parent.filter(
				(child) => child !== change.node,
			);
			record.current = readSectionProperties(liveSiblings);
			const snapshot = change.node.findChild("w:sectPr");
			record.prior = snapshot ? readSectionProperties(snapshot.children) : {};
		}
		if (change.kind === "pPrChange") {
			// Live siblings = the post-edit pPr children; the snapshot's inner
			// <w:pPr> = the prior pPr children.
			const liveSiblings = change.parent.filter(
				(child) => child !== change.node,
			);
			record.current = readParagraphPropsSummary(liveSiblings);
			const snapshot = change.node.findChild("w:pPr");
			record.prior = snapshot
				? readParagraphPropsSummary(snapshot.children)
				: {};
		}
		byId.set(change.id, record);
	}

	const sorted = [...byId.values()].sort(
		(a, b) => trackedChangeIndex(a.id) - trackedChangeIndex(b.id),
	);

	// Tag the two halves of each del+ins replace with a shared `revN` so an agent
	// can accept/reject the logical change in ONE call (`accept --at revN`) instead
	// of the id-renumbering ping-pong of accepting each half separately.
	const { revOf } = revisionGroups(inventory);
	for (const record of sorted) {
		const group = revOf.get(record.id);
		if (group) record.group = group;
	}
	return sorted;
}

/** A record as `track-changes list --json` prints it: the documented
 *  `{ id, kind, author, date, revisionId, blockId, text, group?, prior?,
 *  current? }` shape, without the table-only hints. */
export function listItem({
	contains: _contains,
	paragraphMark: _paragraphMark,
	...item
}: TrackedChangeRecord): Omit<
	TrackedChangeRecord,
	"contains" | "paragraphMark"
> {
	return item;
}

/** After a SUBSET accept/reject/apply, re-read the saved file and render what
 *  remains (with its renumbered handles) so the next call addresses live ids
 *  rather than a stale guess — the contract-finalize death-spiral antidote.
 *  Returns "" when nothing remains or the file can't be re-opened (the mutation
 *  already succeeded and was acked; the advisory is best-effort). */
export async function remainingTrackedChangesBlock(
	path: string,
	verb: string,
): Promise<string> {
	let document: Document;
	try {
		document = await Document.open(path);
	} catch {
		return "";
	}
	const remaining = collectTrackedChangeRecords(document);
	if (remaining.length === 0) return "";
	return `\nRemaining (ids renumbered after this ${verb}):\n\n${renderTrackedChangeTable(remaining)}`;
}

function trackedChangeIndex(id: string): number {
	const match = id.match(/^tc(\d+)$/);
	return match?.[1] ? Number(match[1]) : 0;
}

/** Summarize the direct paragraph properties in a `<w:pPr>` children array
 *  (style/alignment/spacing/indent) for the `pPrChange` prior/current enrichment.
 *  Values are raw twips, matching `read --ast`. */
function readParagraphPropsSummary(children: XmlNode[]): ParagraphPropsSummary {
	const out: ParagraphPropsSummary = {};
	const style = children
		.find((child) => child.tag === "w:pStyle")
		?.getAttribute("w:val");
	if (style) out.style = style;
	const alignment = children
		.find((child) => child.tag === "w:jc")
		?.getAttribute("w:val");
	if (alignment) out.alignment = alignment;

	const spacingNode = children.find((child) => child.tag === "w:spacing");
	if (spacingNode) {
		const spacing: Record<string, string | number> = {};
		for (const attr of ["w:before", "w:after", "w:line"]) {
			const value = Number(spacingNode.getAttribute(attr));
			if (Number.isFinite(value)) spacing[attr.slice(2)] = value;
		}
		const lineRule = spacingNode.getAttribute("w:lineRule");
		if (lineRule) spacing.lineRule = lineRule;
		if (Object.keys(spacing).length > 0) out.spacing = spacing;
	}

	const indentNode = children.find((child) => child.tag === "w:ind");
	if (indentNode) {
		const indent: Record<string, number> = {};
		// `left`/`right` honor the strict/transitional logical attrs (w:start/w:end)
		// the AST reader (core/ast/read.ts) falls back to, so this summary matches
		// `read --ast` on externally-authored docs that use them.
		const slots: [string, string[]][] = [
			["left", ["w:left", "w:start"]],
			["right", ["w:right", "w:end"]],
			["firstLine", ["w:firstLine"]],
			["hanging", ["w:hanging"]],
		];
		for (const [key, attrs] of slots) {
			const raw = attrs
				.map((attr) => indentNode.getAttribute(attr))
				.find((value) => value !== undefined);
			if (raw === undefined) continue;
			const value = Number(raw);
			if (Number.isFinite(value)) indent[key] = value;
		}
		if (Object.keys(indent).length > 0) out.indent = indent;
	}
	return out;
}
