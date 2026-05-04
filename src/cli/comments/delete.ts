import { type DocView, saveDocView } from "@core";
import { XmlNode } from "@core/parser";
import { parseArgs } from "util";
import { EXIT, fail, openOrFail, respond, writeStdout } from "../respond";
import {
	findCommentByNumericId,
	findCommentParaId,
	removeCommentMarkers,
} from "./helpers";
import { pushTrashEntry } from "./trash";

const HELP = `docx comments delete — remove a comment

Usage:
  docx comments delete FILE --id cN [options]

Required:
  --id ID           Comment id (e.g., c0)

Optional:
  --dry-run         Print what would be removed; do not write the file
  -h, --help        Show this help

The deleted comment is journaled to <dir>/.docx-cli/trash.json so it can
be brought back via "docx comments restore".

Examples:
  docx comments delete doc.docx --id c2
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				id: { type: "string" },
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

	const idInput = parsed.values.id as string | undefined;
	if (!idInput) return fail("USAGE", "Missing --id COMMENT_ID", HELP);
	const numericId = idInput.startsWith("c") ? idInput.slice(1) : idInput;
	const commentId = `c${numericId}`;

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const commentReference = findCommentByNumericId(view, numericId);
	if (!commentReference) {
		return fail("COMMENT_NOT_FOUND", `Comment not found: ${commentId}`);
	}

	const anchor = view.doc.comments.find((c) => c.id === commentId)?.anchor;
	if (!anchor) {
		return fail("COMMENT_NOT_FOUND", `Anchor not found for ${commentId}`);
	}

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "comments.delete",
			dryRun: true,
			path,
			commentId,
		});
		return EXIT.OK;
	}

	const commentXml = XmlNode.serialize([commentReference.node]);
	const paraId = findCommentParaId(view, commentId);
	const extXml = paraId ? extractExtEntryXml(view, paraId) : null;

	await pushTrashEntry(path, {
		file: path.split("/").pop() ?? path,
		deletedAt: new Date().toISOString(),
		commentId,
		anchor,
		commentXml,
		extXml,
	});

	const commentIndex = commentReference.parent.indexOf(commentReference.node);
	if (commentIndex !== -1) commentReference.parent.splice(commentIndex, 1);

	if (paraId && view.commentsExtTree) {
		const extRoot = XmlNode.findRoot(view.commentsExtTree, "w15:commentsEx");
		if (extRoot) {
			extRoot.children = extRoot.children.filter(
				(child) =>
					!(
						child.tag === "w15:commentEx" &&
						child.getAttribute("w15:paraId") === paraId
					),
			);
		}
	}

	removeCommentMarkers(view.documentTree, numericId);

	await saveDocView(view);

	await respond({
		ok: true,
		operation: "comments.delete",
		path,
		commentId,
	});
	return EXIT.OK;
}

function extractExtEntryXml(view: DocView, paraId: string): string | null {
	if (!view.commentsExtTree) return null;
	const root = XmlNode.findRoot(view.commentsExtTree, "w15:commentsEx");
	if (!root) return null;
	for (const child of root.children) {
		if (
			child.tag === "w15:commentEx" &&
			child.getAttribute("w15:paraId") === paraId
		) {
			return XmlNode.serialize([child]);
		}
	}
	return null;
}
