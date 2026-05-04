import { addHyperlinkRelationship, saveDocView } from "@core";
import { XmlNode } from "@core/parser";
import { parseArgs } from "util";
import { EXIT, fail, openOrFail, respond, writeStdout } from "../respond";

const HELP = `docx hyperlinks replace — change a hyperlink's URL

Usage:
  docx hyperlinks replace FILE --at LINK_ID --with URL [options]

Required:
  --at LINK_ID      Existing hyperlink to update (e.g., link0)
  --with URL        New target URL

Optional:
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -h, --help        Show this help

Replaces only the targeted hyperlink. If multiple hyperlinks shared the same
underlying relationship, a new relationship is allocated so the others are
unaffected.

Examples:
  docx hyperlinks replace doc.docx --at link0 --with https://example.com
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				at: { type: "string" },
				with: { type: "string" },
				output: { type: "string", short: "o" },
				"dry-run": { type: "boolean" },
				help: { type: "boolean", short: "h" },
			},
		});
	} catch (parseError) {
		const message =
			parseError instanceof Error ? parseError.message : String(parseError);
		return fail("USAGE", message, HELP);
	}

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const targetId = parsed.values.at as string | undefined;
	if (!targetId) return fail("USAGE", "Missing --at LINK_ID", HELP);

	const newUrl = parsed.values.with as string | undefined;
	if (!newUrl) return fail("USAGE", "Missing --with URL", HELP);

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const reference = view.hyperlinkById.get(targetId);
	if (!reference) {
		return fail("HYPERLINK_NOT_FOUND", `Hyperlink not found: ${targetId}`);
	}

	const relationships = XmlNode.findRoot(
		view.relationshipsTree,
		"Relationships",
	);
	if (!relationships) {
		return fail("UNHANDLED", "Missing <Relationships> root in document rels");
	}

	const existingId = reference.relationshipId;
	const oldUrl = existingId
		? view.hyperlinksByRelationshipId.get(existingId)?.url
		: undefined;
	const sharedCount = existingId
		? countHyperlinkUsages(view.documentTree, existingId)
		: 0;
	const willAllocateNew = !existingId || sharedCount > 1;

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "hyperlinks.replace",
			dryRun: true,
			path,
			hyperlinkId: targetId,
			from: oldUrl,
			to: newUrl,
			sharedRelationship: sharedCount > 1,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	if (willAllocateNew) {
		const newRelationshipId = addHyperlinkRelationship(relationships, newUrl);
		reference.node.setAttribute("r:id", newRelationshipId);
		reference.relationshipId = newRelationshipId;
		view.hyperlinksByRelationshipId.set(newRelationshipId, { url: newUrl });
	} else if (existingId) {
		updateRelationshipTarget(relationships, existingId, newUrl);
		view.hyperlinksByRelationshipId.set(existingId, { url: newUrl });
	}

	await saveDocView(view, outputPath);

	await respond({
		ok: true,
		operation: "hyperlinks.replace",
		path: outputPath ?? path,
		hyperlinkId: targetId,
		from: oldUrl,
		to: newUrl,
	});
	return EXIT.OK;
}

function updateRelationshipTarget(
	relationships: XmlNode,
	relationshipId: string,
	newTarget: string,
): void {
	for (const child of relationships.children) {
		if (child.tag !== "Relationship") continue;
		if (child.getAttribute("Id") === relationshipId) {
			child.setAttribute("Target", newTarget);
			return;
		}
	}
}

function countHyperlinkUsages(
	documentTree: XmlNode[],
	relationshipId: string,
): number {
	let count = 0;
	for (const root of documentTree) {
		count += countInNode(root, relationshipId);
	}
	return count;
}

function countInNode(node: XmlNode, relationshipId: string): number {
	let count = 0;
	if (
		node.tag === "w:hyperlink" &&
		node.getAttribute("r:id") === relationshipId
	) {
		count++;
	}
	for (const child of node.children) {
		count += countInNode(child, relationshipId);
	}
	return count;
}
