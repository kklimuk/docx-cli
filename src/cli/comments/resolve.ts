import { Comments, CommentsError } from "@core";
import { normalizeAndDedupCommentIds, readJsonlIds } from "../parse-helpers";
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

const HELP = `docx comments resolve — mark one or more comments resolved

Usage:
  docx comments resolve FILE --id cN [--id cM ...] [options]
  docx comments resolve FILE --batch FILE.jsonl [options]
  docx comments resolve FILE --batch -    [options]   # JSONL from stdin

Anchor (one required, mutually exclusive):
  --id ID             Comment id (e.g., c0). Repeat for multiple ids:
                      --id c1 --id c3 --id c5. All ids are validated against
                      the pre-mutation tree, so the batch is atomic.
  --batch PATH        JSONL with one {"id": "cN"} per line. Use - for stdin.

Optional:
  --unset             Mark unresolved instead of resolved (applies to all
                      ids in the batch)
  -o, --output PATH   Write to PATH instead of overwriting FILE
  --dry-run           Print what would change; do not write the file
  -v, --verbose       Print the success ack JSON (default: silent on success
                      for single; batch always prints the affected ids)
  -h, --help          Show this help

Examples:
  docx comments resolve doc.docx --id c2
  docx comments resolve doc.docx --id c1 --id c3 --unset
  docx comments resolve doc.docx --batch resolutions.jsonl
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			id: { type: "string", multiple: true },
			batch: { type: "string" },
			unset: { type: "boolean" },
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

	const idsRaw = (parsed.values.id as string[] | undefined) ?? [];
	const batchInput = parsed.values.batch as string | undefined;
	const resolved = !parsed.values.unset;

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
		// Mirror the lens's pre-validation so a stale id surfaces in dry-run.
		const view = document.comments;
		for (const commentId of ordered) {
			const numericId = commentId.slice(1);
			if (!view?.findById(numericId)) {
				return fail("COMMENT_NOT_FOUND", `Comment not found: ${commentId}`);
			}
		}
		await respond({
			ok: true,
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

	if (batchInput || ordered.length > 1) {
		await respond({
			ok: true,
			operation: "comments.resolve",
			path: outputPath ?? path,
			resolved,
			batch: ordered.map((commentId) => ({ commentId, resolved })),
		});
	} else {
		const single = ordered[0];
		if (!single) {
			throw new Error("internal: empty single-shot id list");
		}
		await respondAck({
			ok: true,
			operation: "comments.resolve",
			path: outputPath ?? path,
			commentId: single,
			resolved,
		});
	}
	return EXIT.OK;
}
