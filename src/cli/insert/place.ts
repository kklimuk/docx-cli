import {
	type BlockReference,
	type Document,
	Insert,
	InsertError,
	type InsertSpec,
	type XmlNode,
} from "@core";
import type { ParagraphOptions } from "@core/blocks";
import {
	EXIT,
	fail,
	openOrFail,
	renderVerifyHint,
	resolveBlockOrFail,
	resolveTracked,
	respond,
	respondMinted,
} from "../respond";
import type { RawValues } from "./index";

/** Where a new block goes: relative to a user locator (`--after`/`--before`),
 *  or pinned to a document boundary (`--at-start`/`--at-end`) which needs no
 *  existing locator — the boundary resolves against the open document. */
export type TargetPlacement =
	| { mode: "after" | "before"; locator: string }
	| { boundary: "start" | "end" };

/** Everything {@link placeSpec} needs: a pre-built content spec plus where to
 *  put it and how to save. Every insert-family surface (`insert`, `docx code
 *  add`, `docx equations add`, `docx images add`, `docx tables create`) parses
 *  its own flags into this shape, then hands off — so placement, tracked-insert,
 *  splice, save, and minted-locator reporting live in exactly one place. */
export type PlaceSpecOptions = {
	filePath: string;
	placement: TargetPlacement;
	spec: InsertSpec;
	paragraphOptions: ParagraphOptions;
	authorFlag?: string;
	trackFlag: boolean;
	outputPath?: string;
	dryRun: boolean;
};

/** Open the doc, resolve the placement anchor, build the block(s) via the core
 *  Insert lens, splice + save, and report the minted locator(s). The shared
 *  back-end for `insert` and every noun-verb authoring command, so those verbs
 *  differ only in how they PARSE their content spec, not in how it's placed. */
export async function placeSpec(opts: PlaceSpecOptions): Promise<number> {
	const document = await openOrFail(opts.filePath);
	if (typeof document === "number") return document;

	const resolved = await resolvePlacement(document, opts.placement);
	if (typeof resolved === "number") return resolved;
	const { blockRef, mode, locator } = resolved;

	let blocks: XmlNode[];
	try {
		blocks = await new Insert(document).paragraph(
			blockRef,
			opts.spec,
			opts.paragraphOptions,
			{
				placement: mode,
				authorFlag: opts.authorFlag,
				track: resolveTracked(document, opts.trackFlag),
			},
		);
	} catch (error) {
		if (error instanceof InsertError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	return commitInsert(document, blockRef, blocks, opts, mode, locator);
}

/** Parse the placement flags (`--after`/`--before`/`--at-start`/`--at-end`) into
 *  a {@link TargetPlacement}. Shared by `insert` and the noun-verb authoring
 *  commands; each passes its own `help` so a bad placement prints the right
 *  usage. `--at-start`/`--at-end` need no locator (they resolve at open time). */
export async function parseTargetPlacement(
	values: RawValues,
	help?: string,
): Promise<TargetPlacement | number> {
	const after = values.after as string | undefined;
	const before = values.before as string | undefined;
	const atStart = Boolean(values["at-start"]);
	const atEnd = Boolean(values["at-end"]);
	const chosen = [
		after !== undefined ? "--after" : null,
		before !== undefined ? "--before" : null,
		atStart ? "--at-start" : null,
		atEnd ? "--at-end" : null,
	].filter((flag): flag is string => flag !== null);
	if (chosen.length === 0) {
		return fail(
			"USAGE",
			"Missing placement: pass --after, --before, --at-start, or --at-end",
			help,
		);
	}
	if (chosen.length > 1) {
		return fail(
			"USAGE",
			`Pass exactly one placement, got ${chosen.join(" + ")}`,
			help,
		);
	}
	if (atStart) return { boundary: "start" };
	if (atEnd) return { boundary: "end" };
	if (after !== undefined) return { mode: "after", locator: after };
	return { mode: "before", locator: before as string };
}

/** Turn a parsed placement into the concrete anchor the splice needs: a live
 *  `BlockReference`, the side to insert on, and the locator to report. For a
 *  `--at-start`/`--at-end` boundary this resolves against the open document
 *  (first/last content block, or — for an otherwise-empty body — before the
 *  mandatory trailing `<w:sectPr>`). */
async function resolvePlacement(
	document: Document,
	placement: TargetPlacement,
): Promise<
	| { blockRef: BlockReference; mode: "after" | "before"; locator: string }
	| number
> {
	if ("boundary" in placement) {
		return resolveBoundaryAnchor(document, placement.boundary);
	}
	const blockRef = await resolveBlockOrFail(document, placement.locator);
	if (typeof blockRef === "number") return blockRef;
	return { blockRef, mode: placement.mode, locator: placement.locator };
}

/** Resolve `--at-start` / `--at-end` against the document. Anchors only on
 *  TOP-LEVEL content blocks — `<w:p>` / `<w:tbl>` whose parent IS the body's
 *  child list. Filtering by tag alone is wrong: `blockReferences` also holds
 *  table-CELL paragraphs (tag `w:p`), and the reader registers a table's cell
 *  refs BEFORE the table's own `tN` ref, so a tag-only `refs[0]` on a table-first
 *  doc is the first cell paragraph — `--at-start` would splice INSIDE cell (0,0).
 *  Section sentinels (inline + trailing `<w:sectPr>`, which also enumerate as
 *  `sN`) are excluded too, so `--at-end` lands BEFORE the trailing sectPr. */
async function resolveBoundaryAnchor(
	document: Document,
	boundary: "start" | "end",
): Promise<
	| { blockRef: BlockReference; mode: "after" | "before"; locator: string }
	| number
> {
	const bodyChildren = document.body.body.children;
	const refs = [...document.body.blockReferences.entries()].filter(
		([, ref]) =>
			ref.parent === bodyChildren &&
			(ref.node.tag === "w:p" || ref.node.tag === "w:tbl"),
	);
	if (refs.length > 0) {
		const entry = boundary === "start" ? refs[0] : refs[refs.length - 1];
		if (entry) {
			const [locator, blockRef] = entry;
			return {
				blockRef,
				mode: boundary === "start" ? "before" : "after",
				locator,
			};
		}
	}
	// Empty body — only the mandatory trailing <w:sectPr>. Anchor before it so
	// the first inserted block becomes the document's sole content.
	const sectPr = bodyChildren.find((child) => child.tag === "w:sectPr");
	if (!sectPr) {
		return fail(
			"BLOCK_NOT_FOUND",
			"Document body has no blocks to anchor against",
		);
	}
	return {
		blockRef: { node: sectPr, parent: bodyChildren },
		mode: "before",
		locator: "start",
	};
}

/** Splice the built blocks into the document and persist (unless `--dry-run`),
 *  then re-derive and report the inserted block(s)' locators. */
async function commitInsert(
	document: Document,
	blockRef: BlockReference,
	blocks: XmlNode[],
	opts: PlaceSpecOptions,
	mode: "after" | "before",
	anchorLocator: string,
): Promise<number> {
	if (opts.dryRun) {
		await respond({
			operation: "insert",
			dryRun: true,
			path: opts.filePath,
			anchor: anchorLocator,
			placement: mode,
			...(opts.outputPath ? { output: opts.outputPath } : {}),
		});
		return EXIT.OK;
	}

	const targetIndex = blockRef.parent.indexOf(blockRef.node);
	if (targetIndex === -1) {
		return fail(
			"BLOCK_NOT_FOUND",
			"Block reference is stale (parent does not contain it)",
		);
	}
	const insertIndex = mode === "after" ? targetIndex + 1 : targetIndex;
	blockRef.parent.splice(insertIndex, 0, ...blocks);
	await document.save(opts.outputPath);

	// Positional block ids shift after a structural edit, so the agent can't
	// compute where the new block(s) landed. Re-derive ids from the mutated
	// tree and report each inserted block's locator (one per line by default).
	document.reread();
	const insertedNodes = new Set(blocks);
	const locators: string[] = [];
	for (const [blockId, reference] of document.body.blockReferences) {
		if (insertedNodes.has(reference.node)) locators.push(blockId);
	}

	const destination = opts.outputPath ?? opts.filePath;
	await respondMinted(
		locators,
		{
			ok: true,
			operation: "insert",
			path: destination,
			locators,
			anchor: anchorLocator,
			placement: mode,
		},
		isLayoutAffecting(opts.spec) ? renderVerifyHint(destination) : undefined,
	);
	return EXIT.OK;
}

/** Inserts whose result depends on page layout — multi-column sections, page/
 *  column breaks, sized images, fresh tables — need a render to confirm (`read`
 *  shows their text/structure but not how they land on the page). Plain text,
 *  runs, code, and markdown paragraphs reflow normally and don't. */
function isLayoutAffecting(spec: InsertSpec): boolean {
	return (
		spec.kind === "section" ||
		spec.kind === "image" ||
		spec.kind === "table" ||
		spec.kind === "break"
	);
}
