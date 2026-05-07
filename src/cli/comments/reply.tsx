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
import {
	authorInitials,
	CommentBody,
	ensureCommentParaId,
	ensureCommentsExtPart,
	ensureCommentsPart,
	findCommentByNumericId,
	generateParaId,
	nextCommentId,
} from "./helpers";

const HELP = `docx comments reply — reply to an existing comment

Usage:
  docx comments reply FILE --to cN --text TEXT [options]

Required:
  --to ID           Parent comment id (e.g., c0)
  --text TEXT       Reply body

Optional:
  --author NAME     Author name (default: $DOCX_AUTHOR)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would be added; do not write the file
  -v, --verbose     Print the success ack JSON (default: silent on success)
  -h, --help        Show this help

Examples:
  docx comments reply doc.docx --to c0 --text "Good catch" --author "Reviewer"
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				to: { type: "string" },
				text: { type: "string" },
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

	const parentInput = parsed.values.to as string | undefined;
	const text = parsed.values.text as string | undefined;
	if (!parentInput) return fail("USAGE", "Missing --to PARENT_ID", HELP);
	if (!text) return fail("USAGE", "Missing --text TEXT", HELP);

	const parentNumericId = parentInput.startsWith("c")
		? parentInput.slice(1)
		: parentInput;

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const parentReference = findCommentByNumericId(view, parentNumericId);
	if (!parentReference) {
		return fail(
			"COMMENT_NOT_FOUND",
			`Parent comment not found: c${parentNumericId}`,
		);
	}

	const parentParaId = ensureCommentParaId(view, parentInput);
	if (!parentParaId) {
		return fail(
			"COMMENT_NOT_FOUND",
			`Parent comment c${parentNumericId} could not be assigned a w14:paraId.`,
		);
	}

	const author =
		(parsed.values.author as string | undefined) ?? Bun.env.DOCX_AUTHOR ?? "";
	const date = new Date().toISOString();
	const numericId = nextCommentId(view);
	const replyParaId = generateParaId();
	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: "comments.reply",
			dryRun: true,
			path,
			commentId: `c${numericId}`,
			parentId: `c${parentNumericId}`,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	const commentsRoot = ensureCommentsPart(view);
	commentsRoot.children.push(
		<CommentBody
			options={{
				id: numericId,
				author,
				date,
				initials: authorInitials(author),
				paraId: replyParaId,
				text,
			}}
		/>,
	);

	const extRoot = ensureCommentsExtPart(view);
	extRoot.children.push(
		new XmlNode("w15:commentEx", {
			"w15:paraId": replyParaId,
			"w15:paraIdParent": parentParaId,
			"w15:done": "0",
		}),
	);

	await saveDocView(view, outputPath);

	await respondAck({
		ok: true,
		operation: "comments.reply",
		path: outputPath ?? path,
		commentId: `c${numericId}`,
		parentId: `c${parentNumericId}`,
	});
	return EXIT.OK;
}
