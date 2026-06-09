import { Comments, CommentsError, resolveAuthor } from "@core";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondMinted,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx comments reply — reply to an existing comment

Usage:
  docx comments reply FILE --at cN --text TEXT [options]

Required:
  --at cN           Parent comment id (e.g., c0). The "c" prefix is optional.
  --text TEXT       Reply body

Optional:
  --author NAME     Author name (default: $DOCX_AUTHOR, else "Reviewer")
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would be added; do not write the file
  -v, --verbose     Print the full success ack JSON
  -h, --help        Show this help

Output:
  Prints the new reply's comment id (e.g. c4) on success. --verbose prints the
  full ack {ok:true, operation, path, commentId, parentId}. Errors print
  {code, error, hint?} with a nonzero exit.
  Discover existing comment ids with \`docx comments list FILE\`.

Examples:
  docx comments reply doc.docx --at c0 --text "Good catch" --author "Reviewer"
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			text: { type: "string" },
			author: { type: "string" },
			...SAVE_FLAGS,
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

	const parentInput = parsed.values.at as string | undefined;
	const text = parsed.values.text as string | undefined;
	if (!parentInput) return fail("USAGE", "Missing --at cN", HELP);
	if (!text) return fail("USAGE", "Missing --text TEXT", HELP);

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const parentNumericId = parentInput.startsWith("c")
		? parentInput.slice(1)
		: parentInput;
	const author = resolveAuthor(parsed.values.author as string | undefined);
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

	// The reply is itself a new comment with a fresh id the agent can't
	// reconstruct, so print it by default; --verbose upgrades to the full ack.
	await respondMinted([`c${numericId}`], {
		ok: true,
		operation: "comments.reply",
		path: outputPath ?? path,
		commentId: `c${numericId}`,
		parentId: `c${parentNumericId}`,
	});
	return EXIT.OK;
}
