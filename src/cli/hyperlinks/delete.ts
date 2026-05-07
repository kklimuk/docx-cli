import {
	isTrackChangesEnabled,
	resolveAuthor,
	resolveDate,
	saveDocView,
} from "@core";
import { XmlNode } from "@core/parser";
import { parseArgs } from "util";
import {
	emitAuditComment,
	findContainingParagraph,
	findElementOffsetsInParagraph,
} from "../comments/helpers";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	setVerboseAck,
	writeStdout,
} from "../respond";

const HELP = `docx hyperlinks delete — unwrap a hyperlink (keep the text)

Usage:
  docx hyperlinks delete FILE --at LINK_ID [options]

Required:
  --at LINK_ID      Existing hyperlink to remove (e.g., link0)

Optional:
  --author NAME     Author for the audit comment when track-changes is on
                    (default: $DOCX_AUTHOR)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -v, --verbose     Print the success ack JSON (default: silent on success)
  -h, --help        Show this help

The display text stays in place; only the <w:hyperlink> wrapper is removed.
If the underlying relationship is no longer referenced, it is pruned from the
rels file too.

When track-changes is on, an audit comment is anchored to the surviving text
since OOXML has no native tracked-change form for hyperlink removal.

Examples:
  docx hyperlinks delete doc.docx --at link0
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				at: { type: "string" },
				author: { type: "string" },
				output: { type: "string", short: "o" },
				"dry-run": { type: "boolean" },
				verbose: { type: "boolean", short: "v" },
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

	setVerboseAck(Boolean(parsed.values.verbose));

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const targetId = parsed.values.at as string | undefined;
	if (!targetId) return fail("USAGE", "Missing --at LINK_ID", HELP);

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const reference = view.hyperlinkById.get(targetId);
	if (!reference) {
		return fail("HYPERLINK_NOT_FOUND", `Hyperlink not found: ${targetId}`);
	}

	const oldUrl = reference.relationshipId
		? view.hyperlinksByRelationshipId.get(reference.relationshipId)?.url
		: undefined;

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "hyperlinks.delete",
			dryRun: true,
			path,
			hyperlinkId: targetId,
			from: oldUrl,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	const index = reference.parent.indexOf(reference.node);
	if (index === -1) {
		return fail(
			"HYPERLINK_NOT_FOUND",
			`Hyperlink reference is stale (parent does not contain it): ${targetId}`,
		);
	}

	const trackingOn = isTrackChangesEnabled(view);
	const paragraph = trackingOn
		? findContainingParagraph(view.documentTree, reference.node)
		: null;
	const offsets =
		trackingOn && paragraph
			? findElementOffsetsInParagraph(paragraph, reference.node)
			: null;

	reference.parent.splice(index, 1, ...reference.node.children);
	view.hyperlinkById.delete(targetId);

	if (reference.relationshipId) {
		const remaining = countHyperlinkUsages(
			view.documentTree,
			reference.relationshipId,
		);
		if (remaining === 0) {
			pruneRelationship(view.relationshipsTree, reference.relationshipId);
			view.hyperlinksByRelationshipId.delete(reference.relationshipId);
		}
	}

	if (trackingOn && paragraph && offsets) {
		emitAuditComment(
			view,
			{ kind: "span", paragraph, span: offsets },
			{
				body: `[docx-cli] hyperlink removed (was: ${oldUrl ?? "(none)"})`,
				author: resolveAuthor(parsed.values.author as string | undefined),
				date: resolveDate(),
			},
		);
	}

	await saveDocView(view, outputPath);

	await respondAck({
		ok: true,
		operation: "hyperlinks.delete",
		path: outputPath ?? path,
		hyperlinkId: targetId,
		from: oldUrl,
	});
	return EXIT.OK;
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

function pruneRelationship(
	relationshipsTree: XmlNode[],
	relationshipId: string,
): void {
	const relationships = XmlNode.findRoot(relationshipsTree, "Relationships");
	if (!relationships) return;
	relationships.children = relationships.children.filter(
		(child) =>
			!(
				child.tag === "Relationship" &&
				child.getAttribute("Id") === relationshipId
			),
	);
}
