import {
	type BlockReference,
	type CellReference,
	CellTargetError,
	type Document,
	Insert,
	InsertError,
	type InsertSpec,
	parseCellAt,
	resolveCellReference,
	type XmlNode,
} from "@core";
import type { ParagraphOptions } from "@core/blocks";
import {
	applyCellInsertion,
	cellInsertionAnchor,
	ensureCellEndsWithParagraph,
	reusableEmptyCellParagraph,
} from "@core/table";
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
export type BlockTargetPlacement =
	| { mode: "after" | "before"; locator: string }
	| { boundary: "start" | "end" };

export type TargetPlacement =
	| BlockTargetPlacement
	| { mode: "at"; locator: string };

type ResolvedPlacement =
	| {
			kind: "block";
			blockRef: BlockReference;
			mode: "after" | "before";
			requestedPlacement: "after" | "before" | "at";
			locator: string;
	  }
	| {
			kind: "cell";
			cell: CellReference;
			mode: "after" | "before";
			requestedPlacement: "after" | "before" | "at";
			locator: string;
	  };

type ResolvedBlockPlacement = Extract<ResolvedPlacement, { kind: "block" }>;

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
	/** Bare-cell `--at`/`--before`/`--after` is a top-level `docx insert`
	 * affordance. Other commands sharing placeSpec keep their current forms. */
	allowCellTarget?: boolean;
};

/** Open the doc, resolve the placement anchor, build the block(s) via the core
 *  Insert lens, splice + save, and report the minted locator(s). The shared
 *  back-end for `insert` and every noun-verb authoring command, so those verbs
 *  differ only in how they PARSE their content spec, not in how it's placed. */
export async function placeSpec(opts: PlaceSpecOptions): Promise<number> {
	const document = await openOrFail(opts.filePath);
	if (typeof document === "number") return document;

	const resolved = await resolvePlacement(
		document,
		opts.placement,
		opts.allowCellTarget ?? false,
	);
	if (typeof resolved === "number") return resolved;
	const track = resolveTracked(document, opts.trackFlag);
	if (resolved.kind === "cell") {
		const reusable = reusableEmptyCellParagraph(resolved.cell.node);
		const anchorRef = cellInsertionAnchor(
			resolved.cell,
			resolved.mode,
			reusable,
		);
		if (!anchorRef) {
			return fail(
				"TABLE_STRUCTURE",
				`Cell ${resolved.cell.id} has no direct paragraph to inherit insertion formatting from`,
				"Target an explicit paragraph inside the cell, or repair the malformed cell first.",
			);
		}
		let blocks: XmlNode[];
		try {
			blocks = await new Insert(document).paragraph(
				anchorRef,
				opts.spec,
				opts.paragraphOptions,
				{
					placement: resolved.mode,
					authorFlag: opts.authorFlag,
					track,
					reuseAnchorParagraph: reusable !== null,
				},
			);
		} catch (error) {
			if (error instanceof InsertError) {
				return fail(error.code, error.message, error.hint);
			}
			throw error;
		}
		return commitInsert(
			document,
			opts,
			{ locator: resolved.locator, placement: resolved.requestedPlacement },
			async () => {
				const inserted = applyCellInsertion(
					resolved.cell.node,
					blocks,
					resolved.mode,
					reusable,
				);
				ensureCellEndsWithParagraph(resolved.cell.parent);
				return inserted;
			},
		);
	}

	let blocks: XmlNode[];
	try {
		blocks = await new Insert(document).paragraph(
			resolved.blockRef,
			opts.spec,
			opts.paragraphOptions,
			{
				placement: resolved.mode,
				authorFlag: opts.authorFlag,
				track,
			},
		);
	} catch (error) {
		if (error instanceof InsertError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	return commitInsert(
		document,
		opts,
		{ locator: resolved.locator, placement: resolved.requestedPlacement },
		async () => {
			const { blockRef, mode } = resolved;
			const targetIndex = blockRef.parent.indexOf(blockRef.node);
			if (targetIndex === -1) {
				return fail(
					"BLOCK_NOT_FOUND",
					"Block reference is stale (parent does not contain it)",
				);
			}
			const insertIndex = mode === "after" ? targetIndex + 1 : targetIndex;
			blockRef.parent.splice(insertIndex, 0, ...blocks);
			return blocks;
		},
	);
}

/** Parse placement flags (`--at`/`--after`/`--before`/document boundaries) into
 *  a {@link TargetPlacement}. Top-level insert opts into `--at`; shared noun verbs
 *  retain their existing explicit-side forms. Shared by `insert` and noun-verb authoring
 *  commands; each passes its own `help` so a bad placement prints the right
 *  usage. `--at-start`/`--at-end` need no locator (they resolve at open time). */
export function parseTargetPlacement(
	values: RawValues,
	help?: string,
): Promise<BlockTargetPlacement | number>;
export function parseTargetPlacement(
	values: RawValues,
	help: string | undefined,
	options: { allowAt: true },
): Promise<TargetPlacement | number>;
export async function parseTargetPlacement(
	values: RawValues,
	help?: string,
	options: { allowAt?: boolean } = {},
): Promise<TargetPlacement | number> {
	const at = values.at as string | undefined;
	const after = values.after as string | undefined;
	const before = values.before as string | undefined;
	const atStart = Boolean(values["at-start"]);
	const atEnd = Boolean(values["at-end"]);
	const chosen = [
		at !== undefined ? "--at" : null,
		after !== undefined ? "--after" : null,
		before !== undefined ? "--before" : null,
		atStart ? "--at-start" : null,
		atEnd ? "--at-end" : null,
	].filter((flag): flag is string => flag !== null);
	if (chosen.length === 0) {
		return fail(
			"USAGE",
			options.allowAt
				? "Missing placement: pass --at LOCATOR, --after, --before, --at-start, or --at-end"
				: "Missing placement: pass --after, --before, --at-start, or --at-end",
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
	if (at !== undefined) {
		if (!options.allowAt) {
			return fail(
				"USAGE",
				"--at is supported only by top-level `docx insert`",
				help,
			);
		}
		return { mode: "at", locator: at };
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
export function resolvePlacement(
	document: Document,
	placement: BlockTargetPlacement,
): Promise<ResolvedBlockPlacement | number>;
export function resolvePlacement(
	document: Document,
	placement: TargetPlacement,
	allowCellTarget: boolean,
): Promise<ResolvedPlacement | number>;
export async function resolvePlacement(
	document: Document,
	placement: TargetPlacement,
	allowCellTarget = false,
): Promise<ResolvedPlacement | number> {
	if ("boundary" in placement) {
		return resolveBoundaryAnchor(document, placement.boundary);
	}
	if (parseCellAt(placement.locator)) {
		if (!allowCellTarget) {
			return fail(
				"USAGE",
				"Bare table-cell placement is supported only by top-level `docx insert`",
				"Target an explicit cell paragraph (tN:rRcC:pK) for this command.",
			);
		}
		try {
			return {
				kind: "cell",
				cell: resolveCellReference(document, placement.locator),
				mode: insertSide(placement.mode),
				requestedPlacement: placement.mode,
				locator: placement.locator,
			};
		} catch (error) {
			if (error instanceof CellTargetError) {
				return fail(error.code, error.message, error.hint);
			}
			throw error;
		}
	}
	const blockRef = await resolveBlockOrFail(document, placement.locator);
	if (typeof blockRef === "number") return blockRef;
	return {
		kind: "block",
		blockRef,
		mode: insertSide(placement.mode),
		requestedPlacement: placement.mode,
		locator: placement.locator,
	};
}

/** Which side of the anchor the blocks land on. `--at` is the "put it here"
 *  form — after an ordinary block, at a bare cell's end — so it resolves to
 *  `after`; the explicit sides pass through. One rule, one place: `mode` and
 *  `requestedPlacement` ride together on every resolved placement and must never
 *  disagree. */
export function insertSide(
	requested: "after" | "before" | "at",
): "after" | "before" {
	return requested === "before" ? "before" : "after";
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
): Promise<ResolvedBlockPlacement | number> {
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
			const mode = boundary === "start" ? "before" : "after";
			return {
				kind: "block",
				blockRef,
				mode,
				requestedPlacement: mode,
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
		kind: "block",
		blockRef: { node: sectPr, parent: bodyChildren },
		mode: "before",
		requestedPlacement: "before",
		locator: "start",
	};
}

/** Splice the built blocks into the document and persist (unless `--dry-run`),
 *  then re-derive and report the inserted block(s)' locators.
 *
 *  `splice` owns the one step the block-anchored and bare-cell paths don't
 *  share: it performs the tree mutation and returns the nodes now living in the
 *  document (or a `fail()` exit code). Everything around it — the dry-run
 *  payload, the save, the locator re-derivation, the ack, the render hint —
 *  is written once here, so the two paths can't drift. */
async function commitInsert(
	document: Document,
	opts: PlaceSpecOptions,
	anchor: { locator: string; placement: "after" | "before" | "at" },
	splice: () => Promise<XmlNode[] | number>,
): Promise<number> {
	if (opts.dryRun) {
		await respond({
			operation: "insert",
			dryRun: true,
			path: opts.filePath,
			anchor: anchor.locator,
			placement: anchor.placement,
			...(opts.outputPath ? { output: opts.outputPath } : {}),
		});
		return EXIT.OK;
	}

	const inserted = await splice();
	if (typeof inserted === "number") return inserted;
	await document.save(opts.outputPath);

	// Positional block ids shift after a structural edit, so the agent can't
	// compute where the new block(s) landed. Re-derive ids from the mutated
	// tree and report each inserted block's locator (one per line by default).
	document.reread();
	const locators = mintedLocatorsFor(document, new Set(inserted));

	const destination = opts.outputPath ?? opts.filePath;
	await respondMinted(
		locators,
		{
			ok: true,
			operation: "insert",
			path: destination,
			locators,
			anchor: anchor.locator,
			placement: anchor.placement,
		},
		isLayoutAffecting(opts.spec) ? renderVerifyHint(destination) : undefined,
	);
	return EXIT.OK;
}

/** Locators for the freshly-spliced nodes, read off the re-read tree (iteration
 *  is in document order). Call AFTER `document.reread()` — positional ids shift
 *  on every structural edit, so the pre-splice ids can't be reused. */
export function mintedLocatorsFor(
	document: Document,
	insertedNodes: Set<XmlNode>,
): string[] {
	const locators: string[] = [];
	for (const [blockId, reference] of document.body.blockReferences) {
		if (insertedNodes.has(reference.node)) locators.push(blockId);
	}
	return locators;
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
