import { Comments, CommentsError } from "@core";
import { normalizeAndDedupCommentIds, readJsonlIds } from "../parse-helpers";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx comments resolve — mark one or more comments resolved

Usage:
  docx comments resolve FILE --at cN [--at cM ...] [options]
  docx comments resolve FILE --batch FILE.jsonl [options]
  docx comments resolve FILE --batch -    [options]   # JSONL from stdin

Target (one required, mutually exclusive):
  --at cN             Comment id (e.g., c0). Repeat for multiple ids:
                      --at c1 --at c3 --at c5. All ids are validated against
                      the pre-mutation tree, so the batch is atomic. The "c"
                      prefix is optional.
  --batch PATH        JSONL with one {"id": "cN"} per line. Use - for stdin.

Optional:
  --unset             Mark unresolved instead of resolved (applies to all
                      ids in the batch)
  -o, --output PATH   Write to PATH instead of overwriting FILE
  --dry-run           Print what would change; do not write the file
  -v, --verbose       Print the success ack JSON
  -h, --help          Show this help

Output:
  Prints a one-line confirmation on success (exit 0). --verbose prints {ok:true, operation, path,
  resolved, batch:[{commentId, resolved}]}. Errors print {code, error, hint?}
  with a nonzero exit. Discover comment ids with \`docx comments list FILE\`.

Examples:
  docx comments resolve doc.docx --at c2
  docx comments resolve doc.docx --at c1 --at c3 --unset
  docx comments resolve doc.docx --batch resolutions.jsonl
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string", multiple: true },
			batch: { type: "string" },
			unset: { type: "boolean" },
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

	const atValues = (parsed.values.at as string[] | undefined) ?? [];
	const batchInput = parsed.values.batch as string | undefined;
	const resolved = !parsed.values.unset;

	if (atValues.length > 0 && batchInput) {
		return fail("USAGE", "--at and --batch are mutually exclusive", HELP);
	}
	if (atValues.length === 0 && !batchInput) {
		return fail("USAGE", "Specify --at cN (repeatable) or --batch FILE", HELP);
	}

	let rawIds: string[];
	if (batchInput) {
		try {
			rawIds = await readJsonlIds(batchInput);
		} catch (error) {
			const message = error instanceof Error ? error.message : String(error);
			return fail("USAGE", `Failed to read batch: ${message}`);
		}
		if (rawIds.length === 0) {
			return fail("USAGE", "Batch file is empty");
		}
	} else {
		rawIds = atValues;
	}

	const ordered = normalizeAndDedupCommentIds(rawIds);

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		// Mirror the lens's pre-validation so a stale id surfaces in dry-run.
		const view = document.comments;
		for (const commentId of ordered) {
			const numericId = commentId.slice(1);
			if (!view?.findById(numericId)) {
				return fail("COMMENT_NOT_FOUND", `Comment not found: ${commentId}`);
			}
		}
		await respond({
			operation: "comments.resolve",
			dryRun: true,
			path,
			resolved,
			...(outputPath ? { output: outputPath } : {}),
			batch: ordered.map((commentId) => ({ commentId, resolved })),
		});
		return EXIT.OK;
	}

	try {
		new Comments(document).resolve(ordered, resolved);
	} catch (error) {
		if (error instanceof CommentsError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	await document.save(outputPath);

	// resolve mints no new addressable handle (it flips state on existing
	// comments), so it stays silent on success unless --verbose.
	await respondAck({
		ok: true,
		operation: "comments.resolve",
		path: outputPath ?? path,
		resolved,
		batch: ordered.map((commentId) => ({ commentId, resolved })),
	});
	return EXIT.OK;
}
