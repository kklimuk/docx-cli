import {
	type BlockReference,
	type Document,
	isTrailingSectPr,
	removeInlineSectPr,
	resolveAuthor,
	resolveDate,
	TrackChanges,
} from "@core";
import { Comments, findContainingParagraph } from "@core/comments";
import { XmlNode } from "@core/parser";
import {
	applyTrackedRangeDelete,
	applyUntrackedRangeDelete,
	assertParagraphOnlyTrackedRange,
	TrackedRangeConflictError,
} from "@core/track-changes/replace";
import {
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
	resolveBlockRangeOrFail,
	resolveTracked,
	respond,
	respondAck,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";
import { runDeleteBatch } from "./batch";

const HELP = `docx delete — remove a block at a locator

Usage:
  docx delete FILE --at LOCATOR [options]
  docx delete FILE --batch FILE.jsonl [options]   # remove many, one read
  docx delete FILE --batch -          [options]   # read JSONL from stdin

Locator (required, unless --batch):
  --at LOCATOR      What to remove. One of:
                    pN     paragraph (whole block, with all its runs)
                    pN-pM  contiguous paragraph range (delete as a unit)
                    tN     table (entire table)
                    sN     inline section break — strips the <w:sectPr> from
                           its owning paragraph (the paragraph itself stays);
                           rejects the trailing section break (mandatory in OOXML)
                    Discover ids with \`docx read FILE --ast\`. See
                    \`docx info locators\`.

Batch (--batch PATH | -):
  Remove many blocks from one read. Each JSONL line is { "at": LOCATOR } where
  LOCATOR is a whole-block locator (pN, tN, or a cell paragraph tN:rRcC:pK).
  All locators address the document AS READ — they're resolved to live node
  refs before anything is removed, so deleting one never shifts another's id.
  Range (pN-pM), section (sN), equation (eqN), and span (pN:S-E) deletes run
  one at a time, not in a batch. Don't mix --batch with --at.

  --author NAME     Author for tracked changes (default: $DOCX_AUTHOR)
  --track           Record this deletion as a tracked change even when the
                    document's track-changes toggle is off (OFF by default).
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would be removed; do not write the file
  -v, --verbose     Print the success ack JSON (default: a one-line confirmation)
  -h, --help        Show this help

Tracked behavior:
  When tracking is on, paragraph deletion wraps runs in <w:del> and marks
  the paragraph mark as deleted (accept removes the paragraph by merging it
  forward). Section deletion under tracking emits a [docx-cli] audit comment
  on the owning paragraph if it has runs to anchor on; otherwise (sentinel
  paragraphs from "insert --section" have no runs) the mutation is silent.
  delete --at tN under tracking is rejected (tracked table-row deletion is
  not supported).

Output:
  Prints a one-line confirmation on success (exit 0). --verbose prints {ok:true, operation, path,
  locator}. Errors print {code, error, hint?} with a nonzero exit.

Examples:
  docx delete doc.docx --at p3
  docx delete doc.docx --at t0
  docx delete doc.docx --at s2
  docx delete doc.docx --batch drop.jsonl   # {"at":"p26"} per line
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(args, OPTION_SPEC, HELP);
	if (typeof parsed === "number") return parsed;
	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}
	setVerboseAck(Boolean(parsed.values.verbose));

	const filePath = parsed.positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE argument", HELP);

	// `--batch FILE.jsonl`: remove many blocks from one read (resolve-first).
	const batchSource = parsed.values.batch as string | undefined;
	if (batchSource !== undefined) {
		return runDeleteBatch(filePath, batchSource, parsed.values);
	}

	const locator = parsed.values.at as string | undefined;
	if (!locator)
		return fail("USAGE", "Missing --at LOCATOR (or --batch FILE)", HELP);
	const opts: ValidatedOptions = {
		filePath,
		locator,
		authorFlag: parsed.values.author as string | undefined,
		trackFlag: Boolean(parsed.values.track),
		outputPath: parsed.values.output as string | undefined,
		dryRun: Boolean(parsed.values["dry-run"]),
	};

	const document = await openOrFail(opts.filePath);
	if (typeof document === "number") return document;

	// `pN-pM` range delete. Resolves both endpoints, validates same-parent,
	// then either splices (untracked) or wraps content+marks for the canonical
	// Word-shaped tracked-delete (per `/tmp/range-probe/delete-4.docx`).
	if (isBlockRangeLocator(opts.locator)) {
		return commitRangeDelete(document, opts);
	}

	const blockRef = await resolveBlockOrFail(document, opts.locator);
	if (typeof blockRef === "number") return blockRef;

	if (blockRef.node.tag === "w:sectPr") {
		return commitSectionDelete(document, blockRef, opts);
	}
	return commitBlockDelete(document, blockRef, opts);
}

function isBlockRangeLocator(locator: string): boolean {
	return /^p\d+-p\d+$/.test(locator);
}

async function commitRangeDelete(
	document: Document,
	opts: ValidatedOptions,
): Promise<number> {
	const rangeRef = await resolveBlockRangeOrFail(document, opts.locator);
	if (typeof rangeRef === "number") return rangeRef;

	const tracked = resolveTracked(document, opts.trackFlag);
	if (tracked) {
		try {
			assertParagraphOnlyTrackedRange(rangeRef);
		} catch (error) {
			if (error instanceof TrackedRangeConflictError) {
				return fail("TRACKED_CHANGE_CONFLICT", error.message, error.hint);
			}
			throw error;
		}
	}

	if (opts.dryRun) return respondDryRun(opts);

	if (tracked) {
		applyTrackedRangeDelete(
			document,
			rangeRef.parent,
			rangeRef.startIndex,
			rangeRef.endIndex,
			opts.authorFlag,
		);
	} else {
		applyUntrackedRangeDelete(
			rangeRef.parent,
			rangeRef.startIndex,
			rangeRef.endIndex,
		);
	}

	await document.save(opts.outputPath);
	return emitDeleteAck(opts);
}

const OPTION_SPEC = {
	at: { type: "string" },
	batch: { type: "string" },
	author: { type: "string" },
	track: { type: "boolean" },
	...SAVE_FLAGS,
} as const;

type ValidatedOptions = {
	filePath: string;
	locator: string;
	authorFlag?: string;
	trackFlag: boolean;
	outputPath?: string;
	dryRun: boolean;
};

async function commitSectionDelete(
	document: Document,
	blockRef: BlockReference,
	opts: ValidatedOptions,
): Promise<number> {
	const bodyChildren = findBodyChildren(document);
	if (bodyChildren && isTrailingSectPr(bodyChildren, blockRef.parent)) {
		return fail(
			"USAGE",
			"Cannot delete the trailing section break (mandatory in OOXML)",
			"Use `docx edit --at sN --columns 1` to reset its properties instead.",
		);
	}

	if (opts.dryRun) return respondDryRun(opts);

	const trackingOn = resolveTracked(document, opts.trackFlag);
	const owningParagraph = trackingOn
		? findContainingParagraph(document.documentTree, blockRef.node)
		: null;
	const anchorRun =
		owningParagraph?.children.find((child) => child.tag === "w:r") ?? null;

	removeInlineSectPr(blockRef.node, blockRef.parent);

	if (trackingOn && owningParagraph && anchorRun) {
		new Comments(document).addAudit(
			{ kind: "run", paragraph: owningParagraph, run: anchorRun },
			{
				body: `[docx-cli] section break removed (${opts.locator})`,
				author: resolveAuthor(opts.authorFlag),
				date: resolveDate(),
			},
		);
	}

	await document.save(opts.outputPath);
	return emitDeleteAck(opts);
}

async function commitBlockDelete(
	document: Document,
	blockRef: BlockReference,
	opts: ValidatedOptions,
): Promise<number> {
	const targetIndex = blockRef.parent.indexOf(blockRef.node);
	if (targetIndex === -1) {
		return fail(
			"BLOCK_NOT_FOUND",
			"Block reference is stale (parent does not contain it)",
		);
	}

	if (opts.dryRun) return respondDryRun(opts);

	if (resolveTracked(document, opts.trackFlag)) {
		if (blockRef.node.tag !== "w:p") {
			return fail(
				"TRACKED_CHANGE_CONFLICT",
				"Tracked deletion of non-paragraph blocks (e.g., tables) is not supported",
				"Use `docx track-changes off` first, or delete table contents row-by-row.",
			);
		}
		new TrackChanges(document).applyDeletion(blockRef.node, opts.authorFlag);
	} else {
		blockRef.parent.splice(targetIndex, 1);
	}

	await document.save(opts.outputPath);
	return emitDeleteAck(opts);
}

async function respondDryRun(opts: ValidatedOptions): Promise<number> {
	await respond({
		operation: "delete",
		dryRun: true,
		path: opts.filePath,
		locator: opts.locator,
		...(opts.outputPath ? { output: opts.outputPath } : {}),
	});
	return EXIT.OK;
}

async function emitDeleteAck(opts: ValidatedOptions): Promise<number> {
	await respondAck({
		ok: true,
		operation: "delete",
		path: opts.outputPath ?? opts.filePath,
		locator: opts.locator,
	});
	return EXIT.OK;
}

function findBodyChildren(document: Document): XmlNode[] | null {
	const root = XmlNode.findRoot(document.documentTree, "w:document");
	if (!root) return null;
	const body = root.findChild("w:body");
	if (!body) return null;
	return body.children;
}
