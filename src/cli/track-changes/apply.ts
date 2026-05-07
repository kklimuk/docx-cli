import type { TrackedChangeKind } from "@core";
import { saveDocView } from "@core";
import { XmlNode } from "@core/parser";
import { parseArgs } from "util";
import { EXIT, fail, openOrFail, respond, writeStdout } from "../respond";

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
				at: { type: "string" },
				all: { type: "boolean" },
				output: { type: "string", short: "o" },
				"dry-run": { type: "boolean" },
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

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", help);

	const at = parsed.values.at as string | undefined;
	const all = Boolean(parsed.values.all);
	if (at && all) {
		return fail("USAGE", "--at and --all are mutually exclusive", help);
	}
	if (!at && !all) {
		return fail("USAGE", "Specify --at tcN or --all", help);
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
		const found = allChanges.find((change) => change.id === at);
		if (!found) {
			return fail(
				"TRACKED_CHANGE_NOT_FOUND",
				`Tracked change not found: ${at}`,
			);
		}
		targets = [found];
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

	await saveDocView(view, outputPath);

	await respond({
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
			for (const row of block.findChildren("w:tr")) {
				for (const cell of row.findChildren("w:tc")) {
					visitBlocks(cell.children, out, counter);
				}
			}
			continue;
		}
		if (block.tag === "w:sectPr") {
			visitSectPrChange(block, out, counter);
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
	if (target.kind === "sectPrChange") {
		return verb === "accept" ? "delete" : "restore";
	}
	const isAdditive = target.kind === "ins" || target.kind === "moveTo";
	if (verb === "accept") return isAdditive ? "unwrap" : "delete";
	return isAdditive ? "delete" : "unwrap";
}

function applyAccept(target: ChangeFound): void {
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
};

type ChangeAction =
	| "unwrap"
	| "delete"
	| "restore"
	| "merge"
	| "deleteParagraph";

type ChangeRecord = {
	id: string;
	kind: TrackedChangeKind;
	action: ChangeAction;
	author: string;
	date: string;
};
