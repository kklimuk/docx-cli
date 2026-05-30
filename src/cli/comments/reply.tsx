import { Comments, CommentsError } from "@core";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

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
	const parsed = await tryParseArgs(
		args,
		{
			to: { type: "string" },
			text: { type: "string" },
			author: { type: "string" },
			output: { type: "string", short: "o" },
			"dry-run": { type: "boolean" },
			verbose: { type: "boolean", short: "v" },
			help: { type: "boolean", short: "h" },
		},
		HELP,
	);
	if (typeof parsed === "number") return parsed;

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

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const parentNumericId = parentInput.startsWith("c")
		? parentInput.slice(1)
		: parentInput;
	const author =
		(parsed.values.author as string | undefined) ?? Bun.env.DOCX_AUTHOR ?? "";
	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		// Pre-validate so a stale parent reports in dry-run too.
		if (!document.comments?.findById(parentNumericId)) {
			return fail(
				"COMMENT_NOT_FOUND",
				`Parent comment not found: c${parentNumericId}`,
			);
		}
		const nextId = document.comments.nextId();
		await respond({
			ok: true,
			operation: "comments.reply",
			dryRun: true,
			path,
			commentId: `c${nextId}`,
			parentId: `c${parentNumericId}`,
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	let numericId: string;
	try {
		numericId = new Comments(document).reply(parentInput, text, { author });
	} catch (error) {
		if (error instanceof CommentsError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	await document.save(outputPath);

	await respondAck({
		ok: true,
		operation: "comments.reply",
		path: outputPath ?? path,
		commentId: `c${numericId}`,
		parentId: `c${parentNumericId}`,
	});
	return EXIT.OK;
}
