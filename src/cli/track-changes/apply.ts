import type { TrackedChangeKind } from "@core";
import { saveDocView } from "@core";
import type { XmlNode } from "@core/parser";
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
	// stale as we mutate the tree. Pre-order DFS — same id allocation order
	// as the AST reader.
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
		action: actionFor(target.kind, verb),
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
		if (verb === "accept") applyAccept(target.node, target.parent);
		else applyReject(target.node, target.parent);
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
	let counter = 0;
	function walk(children: XmlNode[]): void {
		for (const node of children) {
			const kind = trackedChangeKindForTag(node.tag);
			if (kind) {
				out.push({
					node,
					parent: children,
					kind,
					id: `tc${counter++}`,
					author: node.getAttribute("w:author") ?? "",
					date: node.getAttribute("w:date") ?? "",
				});
			}
			if (node.children.length > 0) walk(node.children);
		}
	}
	walk(tree);
	return out;
}

function trackedChangeKindForTag(tag: string): TrackedChangeKind | null {
	if (tag === "w:ins") return "ins";
	if (tag === "w:del") return "del";
	if (tag === "w:moveFrom") return "moveFrom";
	if (tag === "w:moveTo") return "moveTo";
	return null;
}

/** "Additive" wrappers (ins / moveTo) carry content that shouldn't be in the
 * baseline — accept means "keep this", reject means "throw it out".
 * "Subtractive" wrappers (del / moveFrom) wrap content stored as <w:delText>
 * — accept means "drop it for real", reject means "restore it as plain text". */
function actionFor(
	kind: TrackedChangeKind,
	verb: ApplyVerb,
): "unwrap" | "delete" {
	const isAdditive = kind === "ins" || kind === "moveTo";
	if (verb === "accept") return isAdditive ? "unwrap" : "delete";
	return isAdditive ? "delete" : "unwrap";
}

function applyAccept(node: XmlNode, parent: XmlNode[]): void {
	if (node.tag === "w:ins" || node.tag === "w:moveTo") {
		unwrapNode(node, parent);
		return;
	}
	deleteNode(node, parent);
}

function applyReject(node: XmlNode, parent: XmlNode[]): void {
	if (node.tag === "w:ins" || node.tag === "w:moveTo") {
		deleteNode(node, parent);
		return;
	}
	// del / moveFrom: contents stored as <w:delText>; restore to <w:t> and unwrap.
	renameDelTextToText(node);
	unwrapNode(node, parent);
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
};

type ChangeRecord = {
	id: string;
	kind: TrackedChangeKind;
	action: "unwrap" | "delete";
	author: string;
	date: string;
};
