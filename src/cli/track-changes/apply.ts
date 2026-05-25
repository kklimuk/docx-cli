import type { DocView, TrackedChangeKind } from "@core";
import { saveDocView } from "@core";
import { type NoteKind, noteConfig } from "@core/notes";
import { XmlNode } from "@core/parser";
import {
	acceptCheckboxToggle,
	findCheckboxToggle,
	rejectCheckboxToggle,
} from "@core/task-list";
import { parseArgs } from "util";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	setVerboseAck,
	writeStdout,
} from "../respond";

export async function runApply(
	args: string[],
	verb: ApplyVerb,
	help: string,
): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				at: { type: "string", multiple: true },
				all: { type: "boolean" },
				output: { type: "string", short: "o" },
				"dry-run": { type: "boolean" },
				verbose: { type: "boolean", short: "v" },
				help: { type: "boolean", short: "h" },
			},
		});
	} catch (parseError) {
		const message =
			parseError instanceof Error ? parseError.message : String(parseError);
		return fail("USAGE", message, help);
	}

	if (parsed.values.help) {
		await writeStdout(help);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", help);

	const atRaw = parsed.values.at as string[] | undefined;
	const all = Boolean(parsed.values.all);
	if (atRaw && atRaw.length > 0 && all) {
		return fail("USAGE", "--at and --all are mutually exclusive", help);
	}
	if ((!atRaw || atRaw.length === 0) && !all) {
		return fail("USAGE", "Specify --at tcN (repeatable) or --all", help);
	}

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	// Walk fresh: stored references in view.trackedChangeReferences would go
	// stale as we mutate the tree. The walk order mirrors the AST reader's
	// (run-level trackings first per paragraph, then pPr-level extras —
	// sectPrChange then paragraph-mark) so tcN ids agree with `track-changes
	// list`. Note bodies come AFTER the document body so body-only revisions
	// (footnote edits) get tcN ids that don't shift document-body ids.
	const allChanges = collectTrackedChanges(view);

	let targets: ChangeFound[];
	if (all) {
		targets = allChanges;
	} else {
		// Resolve all requested ids against the SAME pre-mutation walk so that
		// `--at tc1 --at tc2 --at tc3` refers to the nodes the agent saw at
		// `track-changes list` time, even though sibling ids would shift if
		// applied one-by-one. Dedupe so `--at tc1 --at tc1` doesn't try to
		// apply the same node twice.
		const seen = new Set<string>();
		const ordered: string[] = [];
		for (const id of atRaw ?? []) {
			if (seen.has(id)) continue;
			seen.add(id);
			ordered.push(id);
		}
		const byId = new Map(allChanges.map((change) => [change.id, change]));
		targets = [];
		for (const id of ordered) {
			const found = byId.get(id);
			if (!found) {
				return fail(
					"TRACKED_CHANGE_NOT_FOUND",
					`Tracked change not found: ${id}`,
				);
			}
			targets.push(found);
		}
		// Apply in document order regardless of the order on the command line —
		// reverse pre-order traversal happens at apply time below.
		targets.sort((left, right) => {
			const leftIndex = allChanges.indexOf(left);
			const rightIndex = allChanges.indexOf(right);
			return leftIndex - rightIndex;
		});
	}

	const records: ChangeRecord[] = targets.map((target) => ({
		id: target.id,
		kind: target.kind,
		action: actionFor(target, verb),
		author: target.author,
		date: target.date,
	}));

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: `track-changes.${verb}`,
			dryRun: true,
			path,
			...(outputPath ? { output: outputPath } : {}),
			applied: records,
		});
		return EXIT.OK;
	}

	// Snapshot which (kind, noteId) each target touches BEFORE we mutate the
	// tree — once a tracked-delete is applied, its `<w:del>` wrapper is gone
	// and the post-pass can't trace back which notes were affected. The body-
	// side pairing then walks footnotesTree/endnotesTree, GCs orphans, and
	// normalizes any body-side wrappers whose paired reference-side
	// revision was processed.
	const affectedNotes = collectAffectedNotes(targets);

	// Reverse pre-order so a nested ins/del is processed before its parent —
	// keeps stored parent refs valid for the outer node when we get to it.
	for (const target of [...targets].reverse()) {
		if (verb === "accept") applyAccept(target);
		else applyReject(target);
	}

	// Cell removals (accept-cellDel / reject-cellIns) shrink rows without
	// touching <w:tblGrid>; bring the grid back in line with the widest row.
	resyncTableGrids(view.documentTree);

	// Pair body-side note revisions with the reference-side targets we just
	// applied. For each affected (kind, noteId): if no live reference to it
	// remains in document.xml, GC the entire `<w:footnote>`/`<w:endnote>`;
	// otherwise normalize the body's tracking wrappers (unwrap `<w:ins>`,
	// unwrap `<w:del>` with delText→t, drop paragraph-mark del marker).
	// Empirically matches Word — see `/tmp/fn-probe/{add,delete,edit}-*.docx`.
	applyNotePairing(view, affectedNotes);

	await saveDocView(view, outputPath);

	await respondAck({
		ok: true,
		operation: `track-changes.${verb}`,
		path: outputPath ?? path,
		applied: records,
	});
	return EXIT.OK;
}

export type ApplyVerb = "accept" | "reject";

export function collectTrackedChanges(view: DocView): ChangeFound[] {
	const out: ChangeFound[] = [];
	const counter = { value: 0 };

	const documentNode = XmlNode.findRoot(view.documentTree, "w:document");
	const body = documentNode?.findChild("w:body");
	if (body) visitBlocks(body.children, out, counter);

	// A footnote that's been added or deleted under tracking has revisions on
	// BOTH sides: the doc-body reference run (already collected above) AND the
	// note body. The two are paired logically; we hide the body-side from
	// `list`/`apply` so each footnote add/delete surfaces as ONE tcN — the
	// reference-side one. `applyNotePairing` (post-pass) processes the hidden
	// body-side via the footnote id linkage. Body-only revisions (footnote
	// edits — no reference-side wrapper) stay visible: they're standalone
	// `<w:ins>`/`<w:del>` pairs inside the body, and there's no ambiguity.
	const pairedNotes = pairedNoteIdsFromDocBody(out);
	if (view.footnotesTree)
		visitNotePart(view.footnotesTree, "footnote", pairedNotes, out, counter);
	if (view.endnotesTree)
		visitNotePart(view.endnotesTree, "endnote", pairedNotes, out, counter);

	return out;
}

function pairedNoteIdsFromDocBody(docTargets: ChangeFound[]): Set<string> {
	const out = new Set<string>();
	for (const target of docTargets) {
		collectNoteRefs(target.node, out);
	}
	return out;
}

function visitNotePart(
	tree: XmlNode[],
	kind: NoteKind,
	pairedNotes: Set<string>,
	out: ChangeFound[],
	counter: { value: number },
): void {
	const config = noteConfig(kind);
	const root = XmlNode.findRoot(tree, config.rootTag);
	if (!root) return;
	for (const note of root.children) {
		if (note.tag !== config.itemTag) continue;
		// Skip Word's reserved boilerplate (separator / continuationSeparator)
		// — they're never tracked.
		if (note.getAttribute("w:type")) continue;
		const noteId = note.getAttribute("w:id");
		if (noteId && pairedNotes.has(`${kind}:${noteId}`)) continue;
		visitBlocks(note.children, out, counter);
	}
}

function visitBlocks(
	blocks: XmlNode[],
	out: ChangeFound[],
	counter: { value: number },
): void {
	for (const block of blocks) {
		if (block.tag === "w:p") {
			visitParagraph(block, blocks, out, counter);
			continue;
		}
		if (block.tag === "w:tbl") {
			visitTable(block, out, counter);
			continue;
		}
		if (block.tag === "w:sectPr") {
			visitSectPrChange(block, out, counter);
		}
	}
}

/** Visit a table's structural revisions in the same order as the AST reader:
 * the grid revision (in <w:tblGrid>) first, then per row the row marker
 * (<w:trPr><w:ins>/<w:del>), then per cell the cell marker
 * (<w:tcPr><w:cellIns>/<w:cellDel>) followed by the cell's run-level content. */
function visitTable(
	table: XmlNode,
	out: ChangeFound[],
	counter: { value: number },
): void {
	// tblPr revision first (tblPr precedes tblGrid in the tree), then tblGrid.
	const tblPr = table.findChild("w:tblPr");
	const tblPrChange = tblPr?.findChild("w:tblPrChange");
	if (tblPr && tblPrChange) {
		out.push({
			node: tblPrChange,
			parent: tblPr.children,
			kind: "tblPrChange",
			id: `tc${counter.value++}`,
			author: tblPrChange.getAttribute("w:author") ?? "",
			date: tblPrChange.getAttribute("w:date") ?? "",
		});
	}
	const tblGrid = table.findChild("w:tblGrid");
	const gridChange = tblGrid?.findChild("w:tblGridChange");
	if (tblGrid && gridChange) {
		out.push({
			node: gridChange,
			parent: tblGrid.children,
			kind: "tblGridChange",
			id: `tc${counter.value++}`,
			author: gridChange.getAttribute("w:author") ?? "",
			date: gridChange.getAttribute("w:date") ?? "",
		});
	}
	for (const row of table.findChildren("w:tr")) {
		const trPr = row.findChild("w:trPr");
		const rowMarker = trPr?.children.find(
			(child) => child.tag === "w:ins" || child.tag === "w:del",
		);
		if (trPr && rowMarker) {
			out.push({
				node: rowMarker,
				parent: trPr.children,
				kind: rowMarker.tag === "w:ins" ? "rowIns" : "rowDel",
				id: `tc${counter.value++}`,
				author: rowMarker.getAttribute("w:author") ?? "",
				date: rowMarker.getAttribute("w:date") ?? "",
				tableRow: row,
				tableRowParent: table.children,
			});
		}
		for (const cell of row.findChildren("w:tc")) {
			const tcPr = cell.findChild("w:tcPr");
			const cellMarker = tcPr?.children.find(
				(child) => child.tag === "w:cellIns" || child.tag === "w:cellDel",
			);
			if (tcPr && cellMarker) {
				out.push({
					node: cellMarker,
					parent: tcPr.children,
					kind: cellMarker.tag === "w:cellIns" ? "cellIns" : "cellDel",
					id: `tc${counter.value++}`,
					author: cellMarker.getAttribute("w:author") ?? "",
					date: cellMarker.getAttribute("w:date") ?? "",
					tableCell: cell,
					tableCellParent: row.children,
				});
			}
			const tcPrChange = tcPr?.findChild("w:tcPrChange");
			if (tcPr && tcPrChange) {
				out.push({
					node: tcPrChange,
					parent: tcPr.children,
					kind: "tcPrChange",
					id: `tc${counter.value++}`,
					author: tcPrChange.getAttribute("w:author") ?? "",
					date: tcPrChange.getAttribute("w:date") ?? "",
				});
			}
			visitBlocks(cell.children, out, counter);
		}
	}
}

function visitParagraph(
	paragraph: XmlNode,
	paragraphParent: XmlNode[],
	out: ChangeFound[],
	counter: { value: number },
): void {
	visitRunContainer(paragraph, out, counter);
	const pPr = paragraph.findChild("w:pPr");
	if (!pPr) return;
	const sectPr = pPr.findChild("w:sectPr");
	if (sectPr) visitSectPrChange(sectPr, out, counter);
	const rPr = pPr.findChild("w:rPr");
	if (!rPr) return;
	for (const child of rPr.children) {
		const kind = trackedChangeKindForTag(child.tag);
		if (kind !== "ins" && kind !== "del") continue;
		out.push({
			node: child,
			parent: rPr.children,
			kind,
			id: `tc${counter.value++}`,
			author: child.getAttribute("w:author") ?? "",
			date: child.getAttribute("w:date") ?? "",
			paragraph,
			paragraphParent,
		});
	}
}

function visitRunContainer(
	container: XmlNode,
	out: ChangeFound[],
	counter: { value: number },
): void {
	for (const child of container.children) {
		if (child.tag === "w:pPr") continue;
		// Checkbox-toggle SDT: ins+del pair inside <w:sdtContent> surfaces as a
		// single "checkboxToggle" tcN — must match the reader's order in
		// `detectTaskListState` so `list` and `apply --at tcN` agree.
		if (child.tag === "w:sdt") {
			const toggle = findCheckboxToggle(child);
			if (toggle) {
				out.push({
					node: child,
					parent: container.children,
					kind: "checkboxToggle",
					id: `tc${counter.value++}`,
					author: toggle.ins.getAttribute("w:author") ?? "",
					date: toggle.ins.getAttribute("w:date") ?? "",
				});
				continue;
			}
			// SDT with no live toggle (untouched checkbox, or a checkbox being
			// structurally deleted/inserted via `<w:customXmlDelRange*>` markers
			// — which we don't yet enumerate as a single tracked-change kind).
			// Skip the body so we stay aligned with the reader's walker (which
			// also skips SDTs via `skipNodes`). Without this skip, an SDT-
			// internal `<w:del>` (e.g. the deleted glyph of a structural delete)
			// would surface as a phantom tcN with no blockId and break
			// `list`/`apply --at tcN` agreement. The structural delete itself is
			// preserved in the XmlNode tree; `track-changes accept --all` won't
			// honor it yet — see Phase B-2 in src/core/CLAUDE.md.
			continue;
		}
		const kind = trackedChangeKindForTag(child.tag);
		if (
			kind === "ins" ||
			kind === "del" ||
			kind === "moveFrom" ||
			kind === "moveTo"
		) {
			out.push({
				node: child,
				parent: container.children,
				kind,
				id: `tc${counter.value++}`,
				author: child.getAttribute("w:author") ?? "",
				date: child.getAttribute("w:date") ?? "",
			});
			visitRunContainer(child, out, counter);
			continue;
		}
		if (child.children.length > 0) visitRunContainer(child, out, counter);
	}
}

function visitSectPrChange(
	sectPr: XmlNode,
	out: ChangeFound[],
	counter: { value: number },
): void {
	const change = sectPr.findChild("w:sectPrChange");
	if (!change) return;
	out.push({
		node: change,
		parent: sectPr.children,
		kind: "sectPrChange",
		id: `tc${counter.value++}`,
		author: change.getAttribute("w:author") ?? "",
		date: change.getAttribute("w:date") ?? "",
	});
}

function trackedChangeKindForTag(tag: string): TrackedChangeKind | null {
	if (tag === "w:ins") return "ins";
	if (tag === "w:del") return "del";
	if (tag === "w:moveFrom") return "moveFrom";
	if (tag === "w:moveTo") return "moveTo";
	if (tag === "w:sectPrChange") return "sectPrChange";
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
	if (target.kind === "sectPrChange" || target.kind === "tblGridChange") {
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
	if (target.node.tag === "w:sectPrChange") {
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
): void {
	const snapshot = node.findChild(innerTag);
	parent.length = 0;
	if (snapshot) {
		for (const child of snapshot.children) parent.push(child);
	}
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
function applyNotePairing(view: DocView, affected: Set<string>): void {
	if (affected.size === 0) return;
	for (const key of affected) {
		const colon = key.indexOf(":");
		const kind = key.slice(0, colon) as NoteKind;
		const noteId = key.slice(colon + 1);
		const tree = kind === "footnote" ? view.footnotesTree : view.endnotesTree;
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
			view.documentTree,
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

type ChangeFound = {
	node: XmlNode;
	parent: XmlNode[];
	kind: TrackedChangeKind;
	id: string;
	author: string;
	date: string;
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

type ChangeRecord = {
	id: string;
	kind: TrackedChangeKind;
	action: ChangeAction;
	author: string;
	date: string;
};
