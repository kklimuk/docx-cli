import {
	type BlockReference,
	Insert,
	InsertError,
	LocatorResolveError,
	TrackChanges,
} from "@core";
import type { XmlNode } from "@core/parser";
import { readJsonlObjects } from "../parse-helpers";
import {
	EXIT,
	fail,
	openOrFail,
	resolveTracked,
	respond,
	respondMinted,
} from "../respond";
import {
	chooseContentSpec,
	MARKDOWN_INCOMPATIBLE_FLAGS,
	parseParagraphOptions,
	parseTargetPlacement,
	type RawValues,
} from "./index";

/** `docx insert --batch FILE.jsonl`: many inserts from one read. Each JSONL
 *  line mirrors the CLI flags as keys — `{ after | before, <content>, ...opts }`
 *  (e.g. `{"after":"p3","text":"Hi","style":"Heading2"}`). Every anchor is
 *  resolved to a LIVE node ref and all blocks are BUILT before anything is
 *  spliced, so positional ids never shift out from under a later anchor. The
 *  splice phase recomputes each anchor's position fresh and tracks a per-anchor
 *  offset, so stacked inserts (several entries after the same paragraph) land in
 *  entry order. Minted locators are re-derived after one save+reread. */
export async function runInsertBatch(
	filePath: string,
	batchSource: string,
	values: RawValues,
): Promise<number> {
	const conflicting = SINGLE_SHOT_FLAGS.find(
		(flag) => values[flag] !== undefined && values[flag] !== false,
	);
	if (conflicting) {
		return fail(
			"USAGE",
			`--batch reads each insert from the JSONL file; don't also pass --${conflicting} on the CLI`,
			"Put per-entry fields (after/before, text, markdown, style, …) on each JSONL line.",
		);
	}

	const globalAuthor = values.author as string | undefined;
	const trackFlag = Boolean(values.track);
	const outputPath = values.output as string | undefined;
	const dryRun = Boolean(values["dry-run"]);

	let rawEntries: Record<string, unknown>[];
	try {
		rawEntries = await readJsonlObjects(batchSource);
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail("USAGE", `Failed to read batch: ${message}`);
	}
	if (rawEntries.length === 0) return fail("USAGE", "Batch file is empty");

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	const track = resolveTracked(document, trackFlag);
	// One revision-id allocator shared across every entry: blocks are built (and
	// their <w:ins> ids minted) before any splice, so a per-entry allocator would
	// re-scan the same un-mutated tree and mint duplicate w:ids across entries.
	const allocator = track
		? new TrackChanges(document).createAllocator()
		: undefined;

	// Build phase: resolve every anchor to a live ref and build its blocks
	// WITHOUT touching the body. No splice happens here, so every anchor stays
	// valid for the whole phase.
	const built: BuiltEntry[] = [];
	for (let index = 0; index < rawEntries.length; index++) {
		const entry = rawEntries[index];
		if (entry === undefined) continue;
		const entryValues = entryToRawValues(entry);

		const placement = await parseTargetPlacement(entryValues);
		if (typeof placement === "number") return placement;

		const spec = await chooseContentSpec(entryValues);
		if (typeof spec === "number") return spec;

		if (spec.kind === "markdown") {
			const conflict = MARKDOWN_INCOMPATIBLE_FLAGS.find(
				(flag) => entryValues[flag] !== undefined,
			);
			if (conflict) {
				return fail(
					"USAGE",
					`entry ${index}: --${conflict} can't be combined with markdown (the source controls block styling)`,
				);
			}
		}

		const paragraphOptions = await parseParagraphOptions(entryValues);
		if (typeof paragraphOptions === "number") return paragraphOptions;

		let anchorRef: BlockReference;
		try {
			anchorRef = document.body.resolveBlock(placement.locator);
		} catch (error) {
			if (error instanceof LocatorResolveError) {
				return fail("BLOCK_NOT_FOUND", `entry ${index}: ${error.message}`);
			}
			throw error;
		}

		const author =
			typeof entry.author === "string" ? entry.author : globalAuthor;
		let blocks: XmlNode[];
		try {
			blocks = await new Insert(document).paragraph(
				anchorRef,
				spec,
				paragraphOptions,
				{ placement: placement.mode, authorFlag: author, track, allocator },
			);
		} catch (error) {
			if (error instanceof InsertError) {
				return fail(error.code, `entry ${index}: ${error.message}`, error.hint);
			}
			throw error;
		}

		built.push({
			anchorRef,
			mode: placement.mode,
			blocks,
			locator: placement.locator,
		});
	}

	if (dryRun) {
		await respond({
			operation: "insert",
			dryRun: true,
			path: filePath,
			batch: built.map((entry) => ({
				anchor: entry.locator,
				placement: entry.mode,
			})),
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	// Splice phase: recompute each anchor's index fresh (so cross-anchor shifts
	// are absorbed) and track a per-anchor "after" offset (so several inserts
	// after the same paragraph stack in entry order rather than reversing).
	const afterOffset = new Map<XmlNode, number>();
	const insertedNodes = new Set<XmlNode>();
	for (const entry of built) {
		const baseIndex = entry.anchorRef.parent.indexOf(entry.anchorRef.node);
		if (baseIndex === -1) {
			return fail(
				"BLOCK_NOT_FOUND",
				"Anchor reference is stale (parent does not contain it)",
			);
		}
		let insertIndex: number;
		if (entry.mode === "after") {
			const offset = afterOffset.get(entry.anchorRef.node) ?? 0;
			insertIndex = baseIndex + 1 + offset;
			afterOffset.set(entry.anchorRef.node, offset + entry.blocks.length);
		} else {
			insertIndex = baseIndex;
		}
		entry.anchorRef.parent.splice(insertIndex, 0, ...entry.blocks);
		for (const block of entry.blocks) insertedNodes.add(block);
	}

	await document.save(outputPath);

	// Positional ids shifted; re-derive each inserted block's locator from the
	// freshly-read tree (iteration is in document order).
	document.reread();
	const locators: string[] = [];
	for (const [blockId, reference] of document.body.blockReferences) {
		if (insertedNodes.has(reference.node)) locators.push(blockId);
	}

	await respondMinted(locators, {
		ok: true,
		operation: "insert",
		path: outputPath ?? filePath,
		count: built.length,
		locators,
		batch: built.map((entry) => ({
			anchor: entry.locator,
			placement: entry.mode,
		})),
	});
	return EXIT.OK;
}

type BuiltEntry = {
	anchorRef: BlockReference;
	mode: "after" | "before";
	blocks: XmlNode[];
	locator: string;
};

/** CLI flags that have no meaning under `--batch` (each entry carries its own
 *  anchor, content, and per-entry options). Passing any of these alongside
 *  --batch fails fast rather than silently dropping the agent's intent —
 *  every paragraph/run/content sub-flag is listed, not just placement+content,
 *  so e.g. `insert --batch f.jsonl --style Heading1` is a USAGE error (matching
 *  the edit batch). */
const SINGLE_SHOT_FLAGS = [
	"after",
	"before",
	"text",
	"runs",
	"page-break",
	"column-break",
	"section",
	"table",
	"image",
	"code",
	"code-file",
	"equation",
	"markdown",
	"markdown-file",
	"style",
	"alignment",
	"task",
	"list",
	"list-level",
	"color",
	"bold",
	"italic",
	"url",
	"language",
	"alt",
	"width",
	"height",
	"caption",
	"rows",
	"cols",
	"widths",
	"table-width",
	"borders",
	"layout",
	"columns",
	"type",
	"display",
] as const;

/** Coerce a parsed JSONL object into the `RawValues` shape the insert flag
 *  parsers expect: numbers → strings (the sub-parsers `parseInt`/`parseFloat`),
 *  arrays → JSON strings (`runs` is re-parsed by `parseRunsArg`), booleans and
 *  strings pass through. */
function entryToRawValues(entry: Record<string, unknown>): RawValues {
	const out: Record<string, string | boolean | (string | boolean)[]> = {};
	for (const [key, value] of Object.entries(entry)) {
		if (value === undefined || value === null) continue;
		if (typeof value === "boolean") {
			out[key] = value;
			continue;
		}
		if (typeof value === "number") {
			out[key] = String(value);
			continue;
		}
		if (typeof value === "string") {
			out[key] = value;
			continue;
		}
		out[key] = JSON.stringify(value);
	}
	return out as RawValues;
}
