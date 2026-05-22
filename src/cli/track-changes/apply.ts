import type { TrackedChangeKind } from "@core";
import { saveDocView } from "@core";
import { XmlNode } from "@core/parser";
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
	// list`.
	const allChanges = collectTrackedChanges(view.documentTree);

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

	// Reverse pre-order so a nested ins/del is processed before its parent —
	// keeps stored parent refs valid for the outer node when we get to it.
	for (const target of [...targets].reverse()) {
		if (verb === "accept") applyAccept(target);
		else applyReject(target);
	}

	// Cell removals (accept-cellDel / reject-cellIns) shrink rows without
	// touching <w:tblGrid>; bring the grid back in line with the widest row.
	resyncTableGrids(view.documentTree);

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

function collectTrackedChanges(tree: XmlNode[]): ChangeFound[] {
	const out: ChangeFound[] = [];
	const counter = { value: 0 };

	const documentNode = XmlNode.findRoot(tree, "w:document");
	if (!documentNode) return out;
	const body = documentNode.findChild("w:body");
	if (!body) return out;

	visitBlocks(body.children, out, counter);
	return out;
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
	const isAdditive = target.kind === "ins" || target.kind === "moveTo";
	if (verb === "accept") return isAdditive ? "unwrap" : "delete";
	return isAdditive ? "delete" : "unwrap";
}

function applyAccept(target: ChangeFound): void {
	if (applyTableChange(target, "accept")) return;
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
 * preserved; the next paragraph's pPr is dropped. */
function mergeParagraphWithNext(
	paragraph: XmlNode,
	paragraphParent: XmlNode[],
): void {
	const idx = paragraphParent.indexOf(paragraph);
	if (idx === -1) return;
	const next = paragraphParent[idx + 1];
	if (!next || next.tag !== "w:p") return;
	const toMove: XmlNode[] = [];
	for (const child of next.children) {
		if (child.tag === "w:pPr") continue;
		toMove.push(child);
	}
	paragraph.children.push(...toMove);
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
	| "deleteCell";

type ChangeRecord = {
	id: string;
	kind: TrackedChangeKind;
	action: ChangeAction;
	author: string;
	date: string;
};
