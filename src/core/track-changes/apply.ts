import type { Document } from "../ast/document";
import type { TrackedChangeReference } from "../ast/document/body";
import type { TrackedChangeKind } from "../ast/types";
import { type NoteKind, noteConfig } from "../notes";
import { XmlNode } from "../parser";
import { acceptCheckboxToggle, rejectCheckboxToggle } from "../task-list";

export type ApplyVerb = "accept" | "reject";

/** Thrown by `applyTrackedChanges` / `previewTrackedChanges` when an `--at
 * tcN` id doesn't resolve against the fresh pre-mutation walk. */
export class TrackedChangeNotFoundError extends Error {
	constructor(public id: string) {
		super(`Tracked change not found: ${id}`);
		this.name = "TrackedChangeNotFoundError";
	}
}

/** Resolve `target` (explicit tcN ids or "all") against a fresh walk and
 * return the records that WOULD be applied — id, kind, action, author, date —
 * without mutating. Powers the `--dry-run` preview. Throws
 * `TrackedChangeNotFoundError` for an unknown id. */
export function previewTrackedChanges(
	document: Document,
	target: string[] | "all",
	verb: ApplyVerb,
): ChangeRecord[] {
	return resolveTargets(document, target).map((found) =>
		recordFor(found, verb),
	);
}

/** Resolve + apply tracked changes (accept or reject), including the body-side
 * note pairing and table-grid resync post-passes. Returns the applied records.
 * Throws `TrackedChangeNotFoundError` for an unknown id. Caller saves. */
export function applyTrackedChanges(
	document: Document,
	target: string[] | "all",
	verb: ApplyVerb,
): ChangeRecord[] {
	const targets = resolveTargets(document, target);
	return applyResolvedTargets(
		document,
		targets.map((found) => ({ found, verb })),
	);
}

/** Apply a MIXED set of accept/reject decisions in ONE pass, every target
 * resolved against the pre-mutation tree. This is the `track-changes apply`
 * engine: a finalize is "accept these, reject those" in a single atomic call,
 * so neither tcN/revN renumbering (which only bites BETWEEN separate accept and
 * reject invocations) nor a half-finalized intermediate file can occur. An id
 * appearing in both lists is rejected by the caller before we get here. Throws
 * `TrackedChangeNotFoundError` for an unknown id. Caller saves. */
export function applyTrackedDecisions(
	document: Document,
	accepts: string[],
	rejects: string[],
): ChangeRecord[] {
	// Resolve both lists against the SAME fresh pre-mutation walk, then tag each
	// found target with its verb and apply them together in one reverse-preorder
	// traversal — exactly how a single-verb batch already stays renumber-proof.
	const acceptTargets = resolveTargets(document, accepts).map((found) => ({
		found,
		verb: "accept" as ApplyVerb,
	}));
	const rejectTargets = resolveTargets(document, rejects).map((found) => ({
		found,
		verb: "reject" as ApplyVerb,
	}));
	return applyResolvedTargets(document, [...acceptTargets, ...rejectTargets]);
}

/** Shared apply core: take already-resolved (found, verb) pairs, apply them in
 * reverse pre-order (nested wrapper before its parent), then run the table-grid
 * resync and body-side note-pairing post-passes once over the whole set. */
function applyResolvedTargets(
	document: Document,
	decisions: { found: ChangeFound; verb: ApplyVerb }[],
): ChangeRecord[] {
	// Apply in document order: sort by the index of each target in the fresh
	// inventory so a mixed accept/reject set behaves like the single-verb path.
	const order = collectTrackedChanges(document);
	const sorted = [...decisions].sort(
		(left, right) =>
			indexOfFound(order, left.found) - indexOfFound(order, right.found),
	);
	const records = sorted.map(({ found, verb }) => recordFor(found, verb));

	// Snapshot which (kind, noteId) each target touches BEFORE we mutate the
	// tree — once a tracked-delete is applied, its `<w:del>` wrapper is gone
	// and the post-pass can't trace back which notes were affected. The body-
	// side pairing then walks footnotesTree/endnotesTree, GCs orphans, and
	// normalizes any body-side wrappers whose paired reference-side
	// revision was processed.
	const affectedNotes = collectAffectedNotes(sorted.map(({ found }) => found));

	// Reverse pre-order so a nested ins/del is processed before its parent —
	// keeps stored parent refs valid for the outer node when we get to it.
	for (const { found, verb } of [...sorted].reverse()) {
		if (verb === "accept") applyAccept(found);
		else applyReject(found);
	}

	// Cell removals (accept-cellDel / reject-cellIns) shrink rows without
	// touching <w:tblGrid>; bring the grid back in line with the widest row.
	resyncTableGrids(document.documentTree);

	// Pair body-side note revisions with the reference-side targets we just
	// applied. For each affected (kind, noteId): if no live reference to it
	// remains in document.xml, GC the entire `<w:footnote>`/`<w:endnote>`;
	// otherwise normalize the body's tracking wrappers (unwrap `<w:ins>`,
	// unwrap `<w:del>` with delText→t, drop paragraph-mark del marker).
	// Empirically matches Word — see `/tmp/fn-probe/{add,delete,edit}-*.docx`.
	applyNotePairing(document, affectedNotes);

	return records;
}

/** Index of a resolved target within the inventory, by node identity (the same
 * `<w:ins>`/`<w:del>` node), so the mixed-verb set sorts into document order. */
function indexOfFound(inventory: ChangeFound[], found: ChangeFound): number {
	return inventory.findIndex((change) => change.node === found.node);
}

/** Resolve `target` against a fresh walk. "all" yields every change in
 * document order; an id list dedupes, validates (throwing on unknown), and
 * sorts back into document order so a nested ins/del is processed before its
 * parent at apply time. */
function resolveTargets(
	document: Document,
	target: string[] | "all",
): ChangeFound[] {
	const allChanges = collectTrackedChanges(document);
	if (target === "all") return allChanges;

	const seen = new Set<string>();
	const ordered: string[] = [];
	for (const id of target) {
		if (seen.has(id)) continue;
		seen.add(id);
		ordered.push(id);
	}
	const byId = new Map(allChanges.map((change) => [change.id, change]));
	const targets: ChangeFound[] = [];
	for (const id of ordered) {
		const found = byId.get(id);
		if (!found) throw new TrackedChangeNotFoundError(id);
		targets.push(found);
	}
	targets.sort(
		(left, right) => allChanges.indexOf(left) - allChanges.indexOf(right),
	);
	return targets;
}

function recordFor(target: ChangeFound, verb: ApplyVerb): ChangeRecord {
	return {
		id: target.id,
		kind: target.kind,
		action: actionFor(target, verb),
		author: target.author,
		date: target.date,
	};
}

/** The complete tracked-change inventory in document order, sourced entirely
 * from `document.trackedChangeReferences` — the reader's single walk (body then
 * note parts). No re-walk: the map IS the source of truth, so `list` and
 * `accept`/`reject` can never disagree on `tcN` ids. */
export function collectTrackedChanges(document: Document): ChangeFound[] {
	const out: ChangeFound[] = [];
	for (const [id, reference] of document.trackedChangeReferences) {
		const found = changeFoundFromReference(document, id, reference);
		if (found) out.push(found);
	}
	return out;
}

/** Reconstruct a `ChangeFound` from a reader-built map entry. `kind` comes from
 * the entry's explicit `kind` (table-structural / checkboxToggle, whose tags
 * are ambiguous) or else the wrapper tag. author/date read from the metadata
 * element — the inner `<w:ins>` for a checkbox toggle, otherwise the wrapper
 * itself. Table row/cell scope rides on the entry; paragraph-mark scope is
 * recovered from the `blockId → paragraph` link in `Body.blockReferences`. */
function changeFoundFromReference(
	document: Document,
	id: string,
	reference: TrackedChangeReference,
): ChangeFound | null {
	const kind = reference.kind ?? trackedChangeKindForTag(reference.node.tag);
	if (!kind) return null;
	const metadataNode =
		kind === "checkboxToggle"
			? (reference.node.findChild("w:sdtContent")?.findChild("w:ins") ??
				reference.node)
			: reference.node;
	const found: ChangeFound = {
		node: reference.node,
		parent: reference.parent,
		kind,
		id,
		author: metadataNode.getAttribute("w:author") ?? "",
		date: metadataNode.getAttribute("w:date") ?? "",
		blockId: reference.blockId,
		revisionId: metadataNode.getAttribute("w:id") ?? "",
	};
	if (reference.tableRow) found.tableRow = reference.tableRow;
	if (reference.tableRowParent) found.tableRowParent = reference.tableRowParent;
	if (reference.tableCell) found.tableCell = reference.tableCell;
	if (reference.tableCellParent)
		found.tableCellParent = reference.tableCellParent;
	// A run-level <w:ins>/<w:del> whose parent is the owning paragraph's
	// <w:pPr><w:rPr> children is a paragraph-mark tracking — recover the owning
	// paragraph + its parent so accept-del (merge with next) / reject-ins
	// (delete the paragraph) act on the right scope.
	if (kind === "ins" || kind === "del") {
		const paragraphRef = document.body.blockReferences.get(reference.blockId);
		const rPrChildren = paragraphRef?.node
			.findChild("w:pPr")
			?.findChild("w:rPr")?.children;
		if (paragraphRef && rPrChildren === reference.parent) {
			found.paragraph = paragraphRef.node;
			found.paragraphParent = paragraphRef.parent;
		}
	}
	return found;
}

function trackedChangeKindForTag(tag: string): TrackedChangeKind | null {
	if (tag === "w:ins") return "ins";
	if (tag === "w:del") return "del";
	if (tag === "w:moveFrom") return "moveFrom";
	if (tag === "w:moveTo") return "moveTo";
	if (tag === "w:sectPrChange") return "sectPrChange";
	if (tag === "w:pPrChange") return "pPrChange";
	return null;
}

/** "Additive" wrappers (ins / moveTo) carry content that shouldn't be in the
 * baseline — accept means "keep this", reject means "throw it out".
 * "Subtractive" wrappers (del / moveFrom) wrap content stored as <w:delText>
 * — accept means "drop it for real", reject means "restore it as plain text".
 * sectPrChange is its own shape: a snapshot of prior section properties
 * embedded inside the sectPr — accept drops the snapshot, reject restores
 * the snapshot's children to the parent sectPr.
 * Paragraph-mark `<w:ins>`/`<w:del>` (a self-closing element inside
 * <w:pPr><w:rPr>) tracks the paragraph break itself. accept-ins / reject-del
 * just remove the marker (paragraph stays); reject-ins removes the whole
 * owning paragraph (the inserted break disappears, content merges forward —
 * which for sentinels means the paragraph simply vanishes); accept-del merges
 * the owning paragraph with the next (per ECMA-376 §17.13.5.4). */
function actionFor(target: ChangeFound, verb: ApplyVerb): ChangeAction {
	if (target.paragraph) {
		if (target.kind === "ins") {
			return verb === "accept" ? "delete" : "deleteParagraph";
		}
		if (target.kind === "del") {
			return verb === "accept" ? "merge" : "delete";
		}
	}
	if (
		target.kind === "sectPrChange" ||
		target.kind === "tblGridChange" ||
		target.kind === "pPrChange"
	) {
		return verb === "accept" ? "delete" : "restore";
	}
	if (target.kind === "rowIns")
		return verb === "accept" ? "stripMarker" : "deleteRow";
	if (target.kind === "rowDel")
		return verb === "accept" ? "deleteRow" : "stripMarker";
	if (target.kind === "cellIns") {
		return verb === "accept" ? "stripMarker" : "deleteCell";
	}
	if (target.kind === "cellDel") {
		return verb === "accept" ? "deleteCell" : "stripMarker";
	}
	if (target.kind === "checkboxToggle") {
		return verb === "accept" ? "applyToggle" : "revertToggle";
	}
	const isAdditive = target.kind === "ins" || target.kind === "moveTo";
	if (verb === "accept") return isAdditive ? "unwrap" : "delete";
	return isAdditive ? "delete" : "unwrap";
}

function applyAccept(target: ChangeFound): void {
	if (applyTableChange(target, "accept")) return;
	if (target.kind === "checkboxToggle") {
		acceptCheckboxToggle(target.node);
		return;
	}
	if (target.paragraph && target.paragraphParent) {
		if (target.kind === "ins") {
			deleteNode(target.node, target.parent);
			return;
		}
		if (target.kind === "del") {
			deleteNode(target.node, target.parent);
			mergeParagraphWithNext(target.paragraph, target.paragraphParent);
			return;
		}
	}
	if (target.node.tag === "w:ins" || target.node.tag === "w:moveTo") {
		unwrapNode(target.node, target.parent);
		return;
	}
	if (
		target.node.tag === "w:sectPrChange" ||
		target.node.tag === "w:pPrChange"
	) {
		// Accept a property snapshot: drop the marker, keep the new live props.
		deleteNode(target.node, target.parent);
		return;
	}
	deleteNode(target.node, target.parent);
}

function applyReject(target: ChangeFound): void {
	if (applyTableChange(target, "reject")) return;
	if (target.kind === "checkboxToggle") {
		rejectCheckboxToggle(target.node);
		return;
	}
	if (target.paragraph && target.paragraphParent) {
		if (target.kind === "ins") {
			const idx = target.paragraphParent.indexOf(target.paragraph);
			if (idx !== -1) target.paragraphParent.splice(idx, 1);
			return;
		}
		if (target.kind === "del") {
			deleteNode(target.node, target.parent);
			return;
		}
	}
	if (target.node.tag === "w:ins" || target.node.tag === "w:moveTo") {
		deleteNode(target.node, target.parent);
		return;
	}
	if (target.node.tag === "w:sectPrChange") {
		restoreSectPrSnapshot(target.node, target.parent);
		return;
	}
	if (target.node.tag === "w:pPrChange") {
		// Reject: restore the prior pPr children from the snapshot (the marker,
		// which lived among pPr.children, is dropped with them). Preserve any live
		// `<w:sectPr>`: a section boundary is NOT a tracked paragraph property (it's
		// excluded from the snapshot, and tracked separately via sectPrChange), so
		// the pPrChange snapshot omits it — wiping it on restore would silently
		// drop the section break.
		restorePropertySnapshot(target.node, target.parent, "w:pPr", ["w:sectPr"]);
		return;
	}
	// del / moveFrom: contents stored as <w:delText>; restore to <w:t> and unwrap.
	renameDelTextToText(target.node);
	unwrapNode(target.node, target.parent);
}

/** Apply the table revision kinds (rowIns/rowDel/cellIns/cellDel/tblGridChange/
 * tblPrChange/tcPrChange). Returns true when it handled the target. The
 * `…PrChange` kinds are prior-state snapshots: accept drops the snapshot,
 * reject restores it (same mechanism as sectPrChange). Grid-column counts left
 * inconsistent by cell removals are reconciled afterwards by `resyncTableGrids`. */
function applyTableChange(target: ChangeFound, verb: ApplyVerb): boolean {
	switch (target.kind) {
		case "rowIns":
			// accept: keep row, drop the marker; reject: drop the inserted row.
			if (verb === "accept") deleteNode(target.node, target.parent);
			else removeContainer(target.tableRow, target.tableRowParent);
			return true;
		case "rowDel":
			// accept: drop the row; reject: keep it, drop the marker.
			if (verb === "accept")
				removeContainer(target.tableRow, target.tableRowParent);
			else deleteNode(target.node, target.parent);
			return true;
		case "cellIns":
			if (verb === "accept") deleteNode(target.node, target.parent);
			else removeContainer(target.tableCell, target.tableCellParent);
			return true;
		case "cellDel":
			if (verb === "accept")
				removeContainer(target.tableCell, target.tableCellParent);
			else deleteNode(target.node, target.parent);
			return true;
		case "tblGridChange":
			// accept: keep current grid, drop the snapshot; reject: restore prior.
			if (verb === "accept") deleteNode(target.node, target.parent);
			else restorePropertySnapshot(target.node, target.parent, "w:tblGrid");
			return true;
		case "tblPrChange":
			if (verb === "accept") deleteNode(target.node, target.parent);
			else restorePropertySnapshot(target.node, target.parent, "w:tblPr");
			return true;
		case "tcPrChange":
			if (verb === "accept") deleteNode(target.node, target.parent);
			else restorePropertySnapshot(target.node, target.parent, "w:tcPr");
			return true;
		default:
			return false;
	}
}

function removeContainer(
	node: XmlNode | undefined,
	parent: XmlNode[] | undefined,
): void {
	if (!node || !parent) return;
	const index = parent.indexOf(node);
	if (index !== -1) parent.splice(index, 1);
}

/** Reject a `…PrChange` / `tblGridChange`: replace the live properties element's
 * children with the snapshot's prior children (and drop the change marker,
 * which lived among them). `parent` is the owning tblPr/tblGrid/tcPr children
 * array; `innerTag` is the snapshot wrapper (`w:tblPr`/`w:tblGrid`/`w:tcPr`).
 * Mirrors `restoreSectPrSnapshot`. */
function restorePropertySnapshot(
	node: XmlNode,
	parent: XmlNode[],
	innerTag: string,
	preserveTags: readonly string[] = [],
): void {
	const snapshot = node.findChild(innerTag);
	// Carry forward any live children the snapshot deliberately doesn't model
	// (e.g. a paragraph's `<w:sectPr>` for a pPrChange) so the wholesale replace
	// doesn't drop them. They re-append after the restored snapshot children,
	// which keeps CT_PPr order (sectPr sorts after the snapshot's other props).
	const preserved =
		preserveTags.length > 0
			? parent.filter((child) => preserveTags.includes(child.tag))
			: [];
	parent.length = 0;
	if (snapshot) {
		for (const child of snapshot.children) parent.push(child);
	}
	for (const child of preserved) parent.push(child);
}

/** Reconcile each table's `<w:tblGrid>` column count with its widest row after
 * cell removals (accept-cellDel / reject-cellIns shrink rows but don't touch
 * the grid). Trailing `<w:gridCol>` entries are dropped to match; a short grid
 * is padded by repeating the last column's width. */
function resyncTableGrids(tree: XmlNode[]): void {
	walkTables(tree, (table) => {
		const tblGrid = table.findChild("w:tblGrid");
		if (!tblGrid) return;
		const cols = tblGrid.findChildren("w:gridCol");
		let widest = 0;
		for (const row of table.findChildren("w:tr")) {
			let width = 0;
			for (const cell of row.findChildren("w:tc")) width += cellGridSpan(cell);
			widest = Math.max(widest, width);
		}
		if (widest === 0 || cols.length === widest) return;
		if (cols.length > widest) {
			for (const col of cols.slice(widest)) {
				const index = tblGrid.children.indexOf(col);
				if (index !== -1) tblGrid.children.splice(index, 1);
			}
			return;
		}
		const fillWidth = cols[cols.length - 1]?.getAttribute("w:w") ?? "1440";
		const lastCol = cols[cols.length - 1];
		const insertAt = lastCol
			? tblGrid.children.indexOf(lastCol) + 1
			: tblGrid.children.length;
		const additions = Array.from({ length: widest - cols.length }, () =>
			XmlNode.element("w:gridCol", { "w:w": fillWidth }),
		);
		tblGrid.children.splice(insertAt, 0, ...additions);
	});
}

function walkTables(tree: XmlNode[], visit: (table: XmlNode) => void): void {
	for (const node of tree) {
		if (node.tag === "w:tbl") visit(node);
		if (node.children.length > 0) walkTables(node.children, visit);
	}
}

function cellGridSpan(cell: XmlNode): number {
	const raw = cell
		.findChild("w:tcPr")
		?.findChild("w:gridSpan")
		?.getAttribute("w:val");
	const value = raw ? Number(raw) : Number.NaN;
	return Number.isFinite(value) && value > 1 ? value : 1;
}

/** Accept-del-paragraph-mark: the paragraph break is being deleted, so this
 * paragraph absorbs the next paragraph's runs and the next paragraph
 * vanishes. Per ECMA-376 §17.13.5.4. The current paragraph's pPr is
 * preserved; the next paragraph's pPr is dropped — EXCEPT for `<w:sectPr>`,
 * which represents the section ending at that paragraph break. The merged
 * paragraph is the new end of that section, so its sectPr is lifted onto
 * the current paragraph's pPr (creating one if needed). Without this,
 * range edits that span a section boundary would silently lose the section
 * break on accept. */
function mergeParagraphWithNext(
	paragraph: XmlNode,
	paragraphParent: XmlNode[],
): void {
	const idx = paragraphParent.indexOf(paragraph);
	if (idx === -1) return;
	const next = paragraphParent[idx + 1];
	if (!next || next.tag !== "w:p") return;
	const nextSectPr = next.findChild("w:pPr")?.findChild("w:sectPr");
	const toMove: XmlNode[] = [];
	for (const child of next.children) {
		if (child.tag === "w:pPr") continue;
		toMove.push(child);
	}
	paragraph.children.push(...toMove);
	if (nextSectPr) {
		let pPr = paragraph.findChild("w:pPr");
		if (!pPr) {
			pPr = new XmlNode("w:pPr");
			paragraph.children.unshift(pPr);
		}
		const existing = pPr.findChild("w:sectPr");
		if (existing) {
			const existingIdx = pPr.children.indexOf(existing);
			pPr.children.splice(existingIdx, 1, nextSectPr);
		} else {
			pPr.children.push(nextSectPr);
		}
	}
	paragraphParent.splice(idx + 1, 1);
}

/** Restore section properties from a <w:sectPrChange> snapshot. The snapshot
 * lives in `node.findChild("w:sectPr")` and its children are the prior
 * properties of the parent sectPr. We replace the live siblings (parent
 * array, which is sectPr.children) with the snapshot's children, then drop
 * the change marker itself. */
function restoreSectPrSnapshot(node: XmlNode, parent: XmlNode[]): void {
	const snapshot = node.findChild("w:sectPr");
	parent.length = 0;
	if (snapshot) {
		for (const child of snapshot.children) parent.push(child);
	}
}

function unwrapNode(node: XmlNode, parent: XmlNode[]): void {
	const index = parent.indexOf(node);
	if (index === -1) return;
	parent.splice(index, 1, ...node.children);
}

function deleteNode(node: XmlNode, parent: XmlNode[]): void {
	const index = parent.indexOf(node);
	if (index === -1) return;
	parent.splice(index, 1);
}

function renameDelTextToText(node: XmlNode): void {
	if (node.tag === "w:delText") node.tag = "w:t";
	for (const child of node.children) renameDelTextToText(child);
}

/** Snapshot every (kind, noteId) pair that `targets` touches via a
 *  `<w:footnoteReference>` / `<w:endnoteReference>` descendant. Collected
 *  before applying so the post-pass can locate affected notes even after
 *  their reference-side wrappers are gone. */
function collectAffectedNotes(targets: ChangeFound[]): Set<string> {
	const out = new Set<string>();
	for (const target of targets) {
		collectNoteRefs(target.node, out);
	}
	return out;
}

function collectNoteRefs(node: XmlNode, out: Set<string>): void {
	if (node.tag === noteConfig("footnote").referenceTag) {
		const id = node.getAttribute("w:id");
		if (id) out.add(`footnote:${id}`);
		return;
	}
	if (node.tag === noteConfig("endnote").referenceTag) {
		const id = node.getAttribute("w:id");
		if (id) out.add(`endnote:${id}`);
		return;
	}
	for (const child of node.children) collectNoteRefs(child, out);
}

/** For each affected note: count live references in document.xml; if zero,
 *  GC the body from footnotes.xml/endnotes.xml (matches Word — see
 *  `/tmp/fn-probe/{add-rejected,delete-accepted}.docx`). If any remain,
 *  normalize the body's tracking wrappers — unwrap any `<w:ins>`/`<w:del>`
 *  and drop the paragraph-mark del marker — so the body shape stays consistent
 *  with how the reference-side revision was applied. */
function applyNotePairing(document: Document, affected: Set<string>): void {
	if (affected.size === 0) return;
	for (const key of affected) {
		const colon = key.indexOf(":");
		const kind = key.slice(0, colon) as NoteKind;
		const noteId = key.slice(colon + 1);
		const tree =
			kind === "footnote" ? document.footnotes?.tree : document.endnotes?.tree;
		if (!tree) continue;
		const config = noteConfig(kind);
		const root = XmlNode.findRoot(tree, config.rootTag);
		if (!root) continue;
		const noteIndex = root.children.findIndex(
			(child) =>
				child.tag === config.itemTag && child.getAttribute("w:id") === noteId,
		);
		if (noteIndex === -1) continue;
		const note = root.children[noteIndex];
		if (!note) continue;
		const liveRefs = countLiveReferences(
			document.documentTree,
			config.referenceTag,
			noteId,
		);
		if (liveRefs === 0) {
			root.children.splice(noteIndex, 1);
			continue;
		}
		normalizeNoteBody(note);
	}
}

function countLiveReferences(
	tree: XmlNode[],
	refTag: string,
	noteId: string,
): number {
	let count = 0;
	const visit = (node: XmlNode): void => {
		if (node.tag === refTag && node.getAttribute("w:id") === noteId) {
			count += 1;
			return;
		}
		for (const child of node.children) visit(child);
	};
	for (const root of tree) visit(root);
	return count;
}

/** Unwrap any remaining `<w:ins>`/`<w:del>` body wrappers in the note's
 *  paragraph, with `<w:delText>` → `<w:t>` rename on the del side. Also drop
 *  any paragraph-mark del marker that crept into `<w:pPr><w:rPr>`. Result:
 *  the body looks like an untracked note, ready for accept-side semantics. */
function normalizeNoteBody(note: XmlNode): void {
	for (const paragraph of note.children) {
		if (paragraph.tag !== "w:p") continue;
		paragraph.children = unwrapNoteWrappers(paragraph.children);
		stripParagraphMarkTracking(paragraph);
	}
}

function unwrapNoteWrappers(children: XmlNode[]): XmlNode[] {
	const out: XmlNode[] = [];
	for (const child of children) {
		if (child.tag === "w:ins") {
			out.push(...unwrapNoteWrappers(child.children));
			continue;
		}
		if (child.tag === "w:del") {
			renameDelTextToText(child);
			out.push(...unwrapNoteWrappers(child.children));
			continue;
		}
		out.push(child);
	}
	return out;
}

function stripParagraphMarkTracking(paragraph: XmlNode): void {
	const pPr = paragraph.findChild("w:pPr");
	if (!pPr) return;
	const rPr = pPr.findChild("w:rPr");
	if (!rPr) return;
	rPr.children = rPr.children.filter(
		(child) => child.tag !== "w:ins" && child.tag !== "w:del",
	);
}

export type ChangeFound = {
	node: XmlNode;
	parent: XmlNode[];
	kind: TrackedChangeKind;
	id: string;
	author: string;
	date: string;
	/** The owning block's locator id (`pN`/`tN`/`sN`) — `""` for note-body
	 * revisions, which have no block locator. Carried for `track-changes list`;
	 * accept/reject don't use it. */
	blockId: string;
	/** The wrapper's `w:id` revision id (from the inner `<w:ins>` for a checkbox
	 * toggle). Carried for `track-changes list`; accept/reject don't use it. */
	revisionId: string;
	/** Set when the change is a paragraph-mark `<w:ins>`/`<w:del>` in
	 * `<w:pPr><w:rPr>`. Carries the owning paragraph and its parent array
	 * so accept-del (merge with next paragraph) and reject-ins (delete the
	 * whole owning paragraph) can act on the right scope. */
	paragraph?: XmlNode;
	paragraphParent?: XmlNode[];
	/** Set for rowIns/rowDel: the owning `<w:tr>` and the table's children, so
	 * removing the row acts on the right scope. */
	tableRow?: XmlNode;
	tableRowParent?: XmlNode[];
	/** Set for cellIns/cellDel: the owning `<w:tc>` and the row's children. */
	tableCell?: XmlNode;
	tableCellParent?: XmlNode[];
};

type ChangeAction =
	| "unwrap"
	| "delete"
	| "restore"
	| "merge"
	| "deleteParagraph"
	| "stripMarker"
	| "deleteRow"
	| "deleteCell"
	| "applyToggle"
	| "revertToggle";

export type ChangeRecord = {
	id: string;
	kind: TrackedChangeKind;
	action: ChangeAction;
	author: string;
	date: string;
};
