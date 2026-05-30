import { Comments, CommentsError } from "@core";
import { parseArgs } from "util";
import { normalizeAndDedupCommentIds, readJsonlIds } from "../parse-helpers";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	setVerboseAck,
	writeStdout,
} from "../respond";

const HELP = `docx comments delete — remove one or more comments

Usage:
  docx comments delete FILE --id cN [--id cM ...] [options]
  docx comments delete FILE --batch FILE.jsonl [options]
  docx comments delete FILE --batch -    [options]   # JSONL from stdin

Anchor (one required, mutually exclusive):
  --id ID             Comment id (e.g., c0). Repeat for multiple ids:
                      --id c1 --id c3 --id c5. All ids are validated against
                      the pre-mutation tree, so the batch is atomic.
  --batch PATH        JSONL with one {"id": "cN"} per line. Use - for stdin.

Optional:
  -o, --output PATH   Write to PATH instead of overwriting FILE
  --dry-run           Print what would be removed; do not write the file
  -v, --verbose       Print the success ack JSON (default: silent on success
                      for single; batch always prints the removed ids)
  -h, --help          Show this help

Examples:
  docx comments delete doc.docx --id c2
  docx comments delete doc.docx --id c1 --id c3 --id c7
  docx comments delete doc.docx --batch removals.jsonl
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				id: { type: "string", multiple: true },
				batch: { type: "string" },
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

	const idsRaw = (parsed.values.id as string[] | undefined) ?? [];
	const batchInput = parsed.values.batch as string | undefined;

	if (idsRaw.length > 0 && batchInput) {
		return fail("USAGE", "--id and --batch are mutually exclusive", HELP);
	}
	if (idsRaw.length === 0 && !batchInput) {
		return fail("USAGE", "Specify --id cN (repeatable) or --batch FILE", HELP);
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
		rawIds = idsRaw;
	}

	const ordered = normalizeAndDedupCommentIds(rawIds);

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		// Mirror the lens's pre-validation so a stale id reports the failure
		// even in dry-run mode.
		const view = document.comments;
		for (const commentId of ordered) {
			const numericId = commentId.slice(1);
			if (!view?.findById(numericId)) {
				return fail("COMMENT_NOT_FOUND", `Comment not found: ${commentId}`);
			}
		}
		await respond({
			ok: true,
			operation: "comments.delete",
			dryRun: true,
			path,
			...(outputPath ? { output: outputPath } : {}),
			batch: ordered.map((commentId) => ({ commentId })),
		});
		return EXIT.OK;
	}

	try {
		new Comments(document).delete(ordered);
	} catch (error) {
		if (error instanceof CommentsError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	await document.save(outputPath);

	if (batchInput || ordered.length > 1) {
		await respond({
			ok: true,
			operation: "comments.delete",
			path: outputPath ?? path,
			batch: ordered.map((commentId) => ({ commentId })),
		});
	} else {
		const single = ordered[0];
		if (!single) {
			throw new Error("internal: empty single-shot id list");
		}
		await respondAck({
			ok: true,
			operation: "comments.delete",
			path: outputPath ?? path,
			commentId: single,
		});
	}
	return EXIT.OK;
}
