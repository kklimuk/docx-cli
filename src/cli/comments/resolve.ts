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
	ensureCommentParaId,
	ensureCommentsExtPart,
	findCommentByNumericId,
} from "./helpers";

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
				unset: { type: "boolean" },
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

	const seen = new Set<string>();
	const ordered: string[] = [];
	for (const raw of rawIds) {
		const normalized = raw.startsWith("c") ? raw : `c${raw}`;
		if (seen.has(normalized)) continue;
		seen.add(normalized);
		ordered.push(normalized);
	}

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	// Validate every id up-front so the batch is atomic. We pre-resolve
	// each paraId here too so we don't get half-mutated in the apply loop
	// when a malformed comment body lacks an inner <w:p>: the apply loop
	// would otherwise mutate paraIds for entries 0..N-1 in memory before
	// failing on N. ensureCommentParaId is idempotent — calling it here
	// and again below is safe.
	const paraIdByCommentId = new Map<string, string>();
	for (const commentId of ordered) {
		const numericId = commentId.slice(1);
		if (!findCommentByNumericId(view, numericId)) {
			return fail("COMMENT_NOT_FOUND", `Comment not found: ${commentId}`);
		}
		const paraId = ensureCommentParaId(view, commentId);
		if (!paraId) {
			return fail(
				"COMMENT_NOT_FOUND",
				`Comment ${commentId} could not be assigned a w14:paraId.`,
			);
		}
		paraIdByCommentId.set(commentId, paraId);
	}

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
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

	const extRoot = ensureCommentsExtPart(view);

	for (const commentId of ordered) {
		const paraId = paraIdByCommentId.get(commentId);
		if (!paraId) {
			// Unreachable — pre-validation populated the map for every id.
			throw new Error(`internal: missing paraId for ${commentId}`);
		}
		let entry = extRoot.children.find(
			(child) =>
				child.tag === "w15:commentEx" &&
				child.getAttribute("w15:paraId") === paraId,
		);
		if (!entry) {
			entry = new XmlNode("w15:commentEx", { "w15:paraId": paraId });
			extRoot.children.push(entry);
		}
		if (resolved) entry.setAttribute("w15:done", "1");
		else delete entry.attributes["w15:done"];
	}

	await saveDocView(view, outputPath);

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
