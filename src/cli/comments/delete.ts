import { Comments, removeCommentMarkers } from "@core/comments";
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

type RawEntry = { id?: string };

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

	// Dedupe while preserving the agent's order. Normalize the cN prefix.
	const seen = new Set<string>();
	const ordered: string[] = [];
	for (const raw of rawIds) {
		const normalized = raw.startsWith("c") ? raw : `c${raw}`;
		if (seen.has(normalized)) continue;
		seen.add(normalized);
		ordered.push(normalized);
	}

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const comments = new Comments(document);

	// Validate every id up-front so the batch is atomic — any unknown id
	// aborts before we mutate.
	for (const commentId of ordered) {
		const numericId = commentId.slice(1);
		if (!comments.findById(numericId)) {
			return fail("COMMENT_NOT_FOUND", `Comment not found: ${commentId}`);
		}
	}

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
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

	for (const commentId of ordered) {
		const numericId = commentId.slice(1);
		const commentReference = comments.findById(numericId);
		if (!commentReference) continue; // already validated above

		const paraId = comments.paraIdFor(commentId);

		const commentIndex = commentReference.parent.indexOf(commentReference.node);
		if (commentIndex !== -1) commentReference.parent.splice(commentIndex, 1);

		if (paraId && document.comments?.extendedTree) {
			const extRoot = XmlNode.findRoot(
				document.comments?.extendedTree,
				"w15:commentsEx",
			);
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

		removeCommentMarkers(document.documentTree, numericId);
	}

	await document.save(outputPath);

	if (batchInput || ordered.length > 1) {
		// Batch (or multiple --id): always print the removed ids.
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

async function readJsonlIds(source: string): Promise<string[]> {
	const raw =
		source === "-"
			? await new Response(Bun.stdin.stream()).text()
			: await Bun.file(source).text();
	const ids: string[] = [];
	const lines = raw.split("\n");
	for (let index = 0; index < lines.length; index++) {
		const lineRaw = lines[index];
		if (lineRaw === undefined) continue;
		const line = lineRaw.trim();
		if (line.length === 0) continue;
		let parsed: unknown;
		try {
			parsed = JSON.parse(line);
		} catch (error) {
			const message = error instanceof Error ? error.message : String(error);
			throw new Error(`line ${index + 1}: invalid JSON (${message})`);
		}
		if (
			typeof parsed !== "object" ||
			parsed === null ||
			Array.isArray(parsed)
		) {
			throw new Error(`line ${index + 1}: expected a JSON object`);
		}
		const entry = parsed as RawEntry;
		if (typeof entry.id !== "string" || entry.id.length === 0) {
			throw new Error(`line ${index + 1}: missing "id"`);
		}
		ids.push(entry.id);
	}
	return ids;
}
