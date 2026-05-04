import { openDocView, PkgError, saveDocView } from "@core";
import { XmlNode } from "@core/parser";
import { parseArgs } from "util";
import { EXIT, fail, respond, writeStdout } from "../respond";
import {
	addCommentMarkersToParagraph,
	type CommentSpan,
	ensureCommentsExtPart,
	ensureCommentsPart,
	SpanOutOfRangeError,
} from "./helpers";
import { popTrashEntry } from "./trash";

const HELP = `docx comments restore — undo a recent delete

Usage:
  docx comments restore FILE --id cN [options]

Required:
  --id ID           Comment id to restore (e.g., c0)

Optional:
  --dry-run         Print what would be restored; do not write the file
  -h, --help        Show this help

Pulls the most recent matching entry from <dir>/.docx-cli/trash.json
and re-anchors the comment at its original location.

Examples:
  docx comments restore doc.docx --id c2
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
	const commentId = idInput.startsWith("c") ? idInput : `c${idInput}`;
	const numericId = commentId.slice(1);

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "comments.restore",
			dryRun: true,
			path,
			commentId,
		});
		return EXIT.OK;
	}

	const entry = await popTrashEntry(path, commentId);
	if (!entry) {
		return fail(
			"COMMENT_NOT_FOUND",
			`No trashed entry for ${commentId} in ${path}`,
			"Trash lives at <dir>/.docx-cli/trash.json — make sure it's the same directory.",
		);
	}

	if (entry.anchor.startBlockId !== entry.anchor.endBlockId) {
		return fail(
			"USAGE",
			"Cross-block comment restore is not yet supported",
			"v1 supports restoring single-paragraph comment anchors only.",
		);
	}

	let view: Awaited<ReturnType<typeof openDocView>>;
	try {
		view = await openDocView(path);
	} catch (openError) {
		if (openError instanceof PkgError) {
			if (openError.code === "FILE_NOT_FOUND") {
				return fail("FILE_NOT_FOUND", openError.message);
			}
			if (openError.code === "NOT_A_ZIP") {
				return fail("NOT_A_ZIP", openError.message);
			}
		}
		throw openError;
	}

	const blockId = entry.anchor.startBlockId;
	const block = view.blockReferences.get(blockId);
	if (!block) {
		return fail(
			"BLOCK_NOT_FOUND",
			`Original anchor block ${blockId} no longer exists`,
		);
	}

	const span: CommentSpan = {
		start: entry.anchor.startOffset,
		end: entry.anchor.endOffset,
	};

	try {
		addCommentMarkersToParagraph(block.node, numericId, span);
	} catch (error) {
		if (error instanceof SpanOutOfRangeError) {
			return fail(
				"INVALID_LOCATOR",
				`Saved span no longer fits the block: ${error.message}`,
			);
		}
		throw error;
	}

	const commentNodes = XmlNode.parse(entry.commentXml);
	const commentNode = commentNodes[0];
	if (!commentNode) {
		return fail("USAGE", "Trashed commentXml is empty");
	}
	const commentsRoot = ensureCommentsPart(view);
	commentsRoot.children.push(commentNode);

	if (entry.extXml) {
		const extNodes = XmlNode.parse(entry.extXml);
		const extNode = extNodes[0];
		if (extNode) {
			const extRoot = ensureCommentsExtPart(view);
			extRoot.children.push(extNode);
		}
	}

	await saveDocView(view);

	await respond({
		ok: true,
		operation: "comments.restore",
		path,
		commentId,
	});
	return EXIT.OK;
}
