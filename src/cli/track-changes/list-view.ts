import type { SectionProperties, TrackedChange } from "@core";
import { Document, flattenParagraphs, readSectionProperties } from "@core";
import type { XmlNode } from "@core/parser";
import { TrackChanges } from "@core/track-changes";
import { revisionGroups } from "./groups";

/** Compact prior/current summary for a `pPrChange`: the direct paragraph
 *  properties (raw twips, matching `read --ast`). */
type ParagraphPropsSummary = {
	style?: string;
	alignment?: string;
	spacing?: Record<string, string | number>;
	indent?: Record<string, number>;
};

export type TrackedChangeRecord = TrackedChange & {
	blockId: string;
	text: string;
	prior?: SectionProperties | ParagraphPropsSummary;
	current?: SectionProperties | ParagraphPropsSummary;
	/** `revN` when this change is one half of a del+ins replace pair; absent for
	 *  solo changes. `accept/reject --at revN` acts on both halves at once. */
	group?: string;
};

/** Walk a document's tracked changes into the enriched, `revN`-grouped records
 *  that both `track-changes list` and the post-accept/reject "remaining" view
 *  render. Sorted by tcN (document order). */
export function collectTrackedChangeRecords(
	document: Document,
): TrackedChangeRecord[] {
	const byId = new Map<string, TrackedChangeRecord>();
	for (const paragraph of flattenParagraphs(document.body.blocks)) {
		for (const run of paragraph.runs) {
			if (run.type !== "text" || !run.trackedChange) continue;
			const change = run.trackedChange;
			const existing = byId.get(change.id);
			if (existing) {
				existing.text += run.text;
				continue;
			}
			byId.set(change.id, {
				...change,
				blockId: paragraph.id,
				text: run.text,
			});
		}
	}

	// The AST loop above only sees text-bearing run-level changes. Everything
	// else — empty wrappers (e.g. <w:ins> wrapping only <w:del>), section /
	// table-property revisions, checkbox toggles, and standalone note-body
	// edits — comes from the single tracked-change inventory the reader built
	// (TrackChanges.list reads document.trackedChangeReferences; no re-walk). kind,
	// author, date and revisionId are already resolved on each record.
	for (const change of new TrackChanges(document).list()) {
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
	const { revOf } = revisionGroups(sorted);
	for (const record of sorted) {
		const group = revOf.get(record.id);
		if (group) record.group = group;
	}
	return sorted;
}

/** Text-first view: one LOGICAL change per line, keyed by the handle an agent
 *  passes to `accept`/`reject --at`. A del+ins replace pair collapses onto one
 *  `revN` line (the whole point — it pushes weak agents toward the atomic handle
 *  instead of addressing the two renumbering tcN halves separately). */
export function renderTrackedChangeTable(
	records: TrackedChangeRecord[],
): string {
	if (records.length === 0) return "no tracked changes\n";

	const byGroup = new Map<string, TrackedChangeRecord[]>();
	for (const record of records) {
		if (!record.group) continue;
		const members = byGroup.get(record.group) ?? [];
		members.push(record);
		byGroup.set(record.group, members);
	}

	const rows: { handle: string; verb: string; block: string; desc: string }[] =
		[];
	const consumed = new Set<string>();
	for (const record of records) {
		if (record.group) {
			if (consumed.has(record.group)) continue;
			consumed.add(record.group);
			const members = byGroup.get(record.group) ?? [record];
			const del = members.find((member) => member.kind === "del");
			const ins = members.find((member) => member.kind === "ins");
			rows.push({
				handle: record.group,
				verb: "replace",
				block: record.blockId,
				desc: `${quote(del?.text ?? "")} → ${quote(ins?.text ?? "")}`,
			});
			continue;
		}
		rows.push({
			handle: record.id,
			verb: verbLabel(record.kind),
			block: record.blockId,
			desc: describeSolo(record),
		});
	}

	const handleWidth = Math.max(...rows.map((row) => row.handle.length));
	const verbWidth = Math.max(...rows.map((row) => row.verb.length));
	const blockWidth = Math.max(...rows.map((row) => row.block.length));
	const body = rows
		.map(
			(row) =>
				`  ${row.handle.padEnd(handleWidth)}  ${row.verb.padEnd(verbWidth)}  ${row.block.padEnd(blockWidth)}  ${row.desc}`,
		)
		.join("\n");

	const total = records.length;
	const header = `${total} tracked change${total === 1 ? "" : "s"}${
		rows.length === total ? "" : ` (${rows.length} logical)`
	}:`;
	const footer =
		"Address by the leftmost handle: accept/reject --at <handle> (repeatable).\n" +
		"A revN handle resolves both halves of a replace in one call; --json for full detail.";
	return `${header}\n\n${body}\n\n${footer}\n`;
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

/** Quote a run's text for the table, truncating long content with an ellipsis so
 *  one line stays readable. Empty text renders as a paragraph-mark glyph. */
function quote(text: string, max = 44): string {
	const collapsed = text.replace(/\s+/g, " ").trim();
	if (collapsed === "") return "¶";
	const clipped =
		collapsed.length > max ? `${collapsed.slice(0, max - 1)}…` : collapsed;
	return `"${clipped}"`;
}

/** Friendly verb for a solo change's kind (the `revN` replace rows bypass this). */
function verbLabel(kind: TrackedChange["kind"]): string {
	const labels: Record<string, string> = {
		ins: "insert",
		del: "delete",
		moveFrom: "move-from",
		moveTo: "move-to",
		sectPrChange: "section",
		pPrChange: "format",
		rowIns: "row-insert",
		rowDel: "row-delete",
		cellIns: "cell-insert",
		cellDel: "cell-delete",
		tblGridChange: "grid",
		tblPrChange: "table-fmt",
		tcPrChange: "cell-fmt",
		checkboxToggle: "checkbox",
	};
	return labels[kind] ?? kind;
}

/** Human description for a non-grouped change. Text runs show their content;
 *  property revisions show the changed keys; structural changes a short note. */
function describeSolo(record: TrackedChangeRecord): string {
	// ins/del AND moveFrom/moveTo all carry the affected run text — show it.
	// (moveFrom/moveTo are the only other kinds with non-empty text; every
	// table/property/checkbox kind has text "" and gets a fixed note below.)
	if (
		record.kind === "ins" ||
		record.kind === "del" ||
		record.kind === "moveFrom" ||
		record.kind === "moveTo"
	) {
		return record.text === "" ? "¶ paragraph mark" : quote(record.text);
	}
	if (record.kind === "pPrChange" || record.kind === "sectPrChange") {
		const diff = propDiff(record.prior, record.current);
		return (
			diff || (record.kind === "pPrChange" ? "paragraph format" : "layout")
		);
	}
	const notes: Record<string, string> = {
		rowIns: "row inserted",
		rowDel: "row deleted",
		cellIns: "cell inserted",
		cellDel: "cell deleted",
		tblGridChange: "grid widths",
		tblPrChange: "table properties",
		tcPrChange: "cell properties",
		checkboxToggle: "checkbox toggled",
	};
	return notes[record.kind] ?? "";
}

/** Compact "key →value" summary of the changed leaf properties between a
 *  property revision's prior and current snapshots, so an agent can see WHAT a
 *  `format`/`section` change did without `--json`. Caps at three keys. */
function propDiff(prior: unknown, current: unknown): string {
	const before = flattenProps(prior);
	const after = flattenProps(current);
	const keys = new Set([...before.keys(), ...after.keys()]);
	const parts: string[] = [];
	for (const key of keys) {
		const from = before.get(key);
		const to = after.get(key);
		if (from === to) continue;
		parts.push(`${key} ${from ?? "·"}→${to ?? "·"}`);
	}
	if (parts.length === 0) return "";
	if (parts.length > 3) return `${parts.slice(0, 3).join(", ")}, …`;
	return parts.join(", ");
}

/** Flatten a nested prior/current props object to dotted leaf keys (e.g.
 *  `spacing.line` → 360) for {@link propDiff}. */
function flattenProps(obj: unknown, prefix = ""): Map<string, string | number> {
	const out = new Map<string, string | number>();
	if (obj === null || typeof obj !== "object") return out;
	for (const [key, value] of Object.entries(obj)) {
		const path = prefix ? `${prefix}.${key}` : key;
		if (value !== null && typeof value === "object") {
			for (const [k, v] of flattenProps(value, path)) out.set(k, v);
		} else if (typeof value === "string" || typeof value === "number") {
			out.set(path, value);
		}
	}
	return out;
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
