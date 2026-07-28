import {
	type BlockReference,
	type CellReference,
	CellTargetError,
	Insert,
	InsertError,
	LocatorResolveError,
	parseCellAt,
	resolveCellReference,
	TrackChanges,
} from "@core";
import type { XmlNode } from "@core/parser";
import {
	CellInsertionCursor,
	cellInsertionAnchor,
	type Grid,
	reusableEmptyCellParagraph,
} from "@core/table";
import { readJsonlObjects, rejectBatchOnlyFlags } from "../parse-helpers";
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
	type RawValues,
} from "./index";
import { insertSide, mintedLocatorsFor, parseTargetPlacement } from "./place";

/** `docx insert --batch FILE.jsonl`: many inserts from one read. Each JSONL
 *  line mirrors the CLI flags as keys — `{ at | after | before, <content>, ...opts }`
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
	const conflict = await rejectBatchOnlyFlags(
		values,
		SINGLE_SHOT_FLAGS,
		"insert",
		"Put per-entry fields (after/before, text, markdown, style, …) on each JSONL line.",
	);
	if (conflict !== undefined) return conflict;

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
	// valid for the whole phase. A bare cell and one of its explicit direct blocks
	// cannot share a batch: empty-cell reuse would consume that block anchor.
	const built: BuiltEntry[] = [];
	// The build phase never splices, so one grid per table stays valid across
	// every entry — a many-cell fill of one table builds it once, not per entry.
	const gridCache = new Map<XmlNode, Grid>();
	const reservedEmptyCells = new Set<XmlNode>();
	const cellTargets = new Map<XmlNode[], { index: number; locator: string }>();
	const blockTargets = new Map<XmlNode[], { index: number; locator: string }>();
	for (let index = 0; index < rawEntries.length; index++) {
		const entry = rawEntries[index];
		if (entry === undefined) continue;
		const entryValues = entryToRawValues(entry);

		const placement = await parseTargetPlacement(entryValues, undefined, {
			allowAt: true,
		});
		if (typeof placement === "number") return placement;
		if ("boundary" in placement) {
			return fail(
				"USAGE",
				`entry ${index}: --at-start/--at-end aren't supported in --batch (use --after/--before with a locator)`,
			);
		}

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

		const author =
			typeof entry.author === "string" ? entry.author : globalAuthor;
		const cellTarget = parseCellAt(placement.locator);
		if (cellTarget) {
			let cell: CellReference;
			try {
				cell = resolveCellReference(document, placement.locator, gridCache);
			} catch (error) {
				if (error instanceof CellTargetError) {
					return fail(
						error.code,
						`entry ${index}: ${error.message}`,
						error.hint,
					);
				}
				throw error;
			}
			const explicitTarget = blockTargets.get(cell.parent);
			if (explicitTarget) {
				return mixedCellTargetFailure(index, placement.locator, explicitTarget);
			}
			cellTargets.set(cell.parent, { index, locator: placement.locator });

			const candidate = reusableEmptyCellParagraph(cell.node);
			const reusable =
				candidate && !reservedEmptyCells.has(cell.node) ? candidate : null;
			if (reusable) reservedEmptyCells.add(cell.node);
			const mode = insertSide(placement.mode);
			const anchorRef = cellInsertionAnchor(cell, mode, reusable);
			if (!anchorRef) {
				return fail(
					"TABLE_STRUCTURE",
					`entry ${index}: Cell ${cell.id} has no direct paragraph to inherit insertion formatting from`,
					"Target an explicit paragraph inside the cell, or repair the malformed cell first.",
				);
			}

			let blocks: XmlNode[];
			try {
				blocks = await new Insert(document).paragraph(
					anchorRef,
					spec,
					paragraphOptions,
					{
						placement: mode,
						authorFlag: author,
						track,
						allocator,
						reuseAnchorParagraph: reusable !== null,
					},
				);
			} catch (error) {
				if (error instanceof InsertError) {
					return fail(
						error.code,
						`entry ${index}: ${error.message}`,
						error.hint,
					);
				}
				throw error;
			}
			built.push({
				kind: "cell",
				index,
				cell,
				mode,
				requestedPlacement: placement.mode,
				reuseParagraph: reusable,
				blocks,
				locator: placement.locator,
			});
			continue;
		}

		const mode = insertSide(placement.mode);
		let anchorRef: BlockReference;
		try {
			anchorRef = document.body.resolveBlock(placement.locator);
		} catch (error) {
			if (error instanceof LocatorResolveError) {
				return fail("BLOCK_NOT_FOUND", `entry ${index}: ${error.message}`);
			}
			throw error;
		}
		const bareTarget = cellTargets.get(anchorRef.parent);
		if (bareTarget) {
			return mixedCellTargetFailure(bareTarget.index, bareTarget.locator, {
				index,
				locator: placement.locator,
			});
		}
		blockTargets.set(anchorRef.parent, { index, locator: placement.locator });

		let blocks: XmlNode[];
		try {
			blocks = await new Insert(document).paragraph(
				anchorRef,
				spec,
				paragraphOptions,
				{
					placement: mode,
					authorFlag: author,
					track,
					allocator,
				},
			);
		} catch (error) {
			if (error instanceof InsertError) {
				return fail(error.code, `entry ${index}: ${error.message}`, error.hint);
			}
			throw error;
		}
		built.push({
			kind: "block",
			index,
			anchorRef,
			mode,
			requestedPlacement: placement.mode,
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
				placement: entry.requestedPlacement,
			})),
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	// Splice phase: recompute block-anchor indexes fresh and keep one
	// CellInsertionCursor per cell (it carries the start cursor; the end boundary
	// re-derives). Cursors preserve JSONL order at each cell boundary while all
	// targets still refer to the document as originally read.
	const afterOffset = new Map<XmlNode, number>();
	const cellCursors = new Map<XmlNode, CellInsertionCursor>();
	const insertedNodes = new Set<XmlNode>();
	for (const entry of built) {
		if (entry.kind === "cell") {
			// Scoped to the entry so the failure names the JSONL line that caused
			// it, the way every build-phase rejection does.
			try {
				const cursor =
					cellCursors.get(entry.cell.node) ??
					CellInsertionCursor.open(entry.cell);
				cellCursors.set(entry.cell.node, cursor);
				const inserted = cursor.insert(
					entry.blocks,
					entry.mode,
					entry.reuseParagraph,
				);
				for (const block of inserted) insertedNodes.add(block);
			} catch (error) {
				if (error instanceof CellTargetError) {
					return fail(
						error.code,
						`entry ${entry.index}: ${error.message}`,
						error.hint,
					);
				}
				throw error;
			}
			continue;
		}

		const baseIndex = entry.anchorRef.parent.indexOf(entry.anchorRef.node);
		if (baseIndex === -1) {
			return fail(
				"BLOCK_NOT_FOUND",
				`entry ${entry.index}: Anchor reference is stale (parent does not contain it)`,
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
	for (const cursor of cellCursors.values()) cursor.ensureTerminalParagraph();

	await document.save(outputPath);

	// Positional ids shifted; re-derive each inserted block's locator from the
	// freshly-read tree (iteration is in document order).
	document.reread();
	const locators = mintedLocatorsFor(document, insertedNodes);

	await respondMinted(locators, {
		ok: true,
		operation: "insert",
		path: outputPath ?? filePath,
		count: built.length,
		locators,
		batch: built.map((entry) => ({
			anchor: entry.locator,
			placement: entry.requestedPlacement,
		})),
	});
	return EXIT.OK;
}

function mixedCellTargetFailure(
	cellEntryIndex: number,
	cellLocator: string,
	explicit: { index: number; locator: string },
): Promise<number> {
	return fail(
		"USAGE",
		`entry ${cellEntryIndex}: bare cell ${cellLocator} conflicts with explicit block target ${explicit.locator} in entry ${explicit.index}`,
		"Put bare-cell and explicit CELL:pK operations in separate batches, with a re-read between them.",
	);
}

/** `index` is the 0-based JSONL line the entry came from, carried through the
 *  build phase so a splice-phase failure names the same `entry N` a build-phase
 *  one does. */
type BuiltEntry =
	| {
			kind: "block";
			index: number;
			anchorRef: BlockReference;
			mode: "after" | "before";
			requestedPlacement: "after" | "before" | "at";
			blocks: XmlNode[];
			locator: string;
	  }
	| {
			kind: "cell";
			index: number;
			cell: CellReference;
			mode: "after" | "before";
			requestedPlacement: "after" | "before" | "at";
			reuseParagraph: XmlNode | null;
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
	"at",
	"after",
	"before",
	"at-start",
	"at-end",
	"text",
	"text-file",
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
	"space-before",
	"space-after",
	"line-spacing",
	"indent-left",
	"indent-right",
	"first-line",
	"hanging",
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
