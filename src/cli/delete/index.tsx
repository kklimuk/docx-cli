import {
	type BlockReference,
	convertTextToDelText,
	createRevisionAllocator,
	Del,
	type DocView,
	isTrackChangesEnabled,
	isTrailingSectPr,
	markParagraphMarkAs,
	removeInlineSectPr,
	resolveAuthor,
	resolveDate,
	saveDocView,
	type TrackedMeta,
} from "@core";
import { XmlNode } from "@core/parser";
import { parseArgs } from "util";
import { emitAuditComment, findContainingParagraph } from "../comments/helpers";
import {
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
	respond,
	respondAck,
	setVerboseAck,
	writeStdout,
} from "../respond";

const HELP = `docx delete — remove a block at a locator

Usage:
  docx delete FILE [options]

Locator (required):
  --at LOCATOR      Block to remove
                    pN  paragraph (whole block, with all its runs)
                    tN  table (entire table)
                    sN  inline section break — strips the <w:sectPr> from
                        its owning paragraph (the paragraph itself stays);
                        rejects the trailing section break (mandatory in OOXML)

  --author NAME     Author for tracked changes (default: $DOCX_AUTHOR)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would be removed; do not write the file
  -v, --verbose     Print the success ack JSON (default: silent on success)
  -h, --help        Show this help

Tracked behavior:
  When tracking is on, paragraph deletion wraps runs in <w:del> and marks
  the paragraph mark as deleted (accept removes the paragraph by merging it
  forward). Section deletion under tracking emits a [docx-cli] audit comment
  on the owning paragraph if it has runs to anchor on; otherwise (sentinel
  paragraphs from "insert --section" have no runs) the mutation is silent.
  delete --at tN under tracking is rejected (tracked table-row deletion is
  not supported).

Examples:
  docx delete doc.docx --at p3
  docx delete doc.docx --at t0
  docx delete doc.docx --at s2
`;

export async function run(args: string[]): Promise<number> {
	const opts = await parseAndValidateOptions(args);
	if (typeof opts === "number") return opts;

	const view = await openOrFail(opts.filePath);
	if (typeof view === "number") return view;

	const blockRef = await resolveBlockOrFail(view, opts.locator);
	if (typeof blockRef === "number") return blockRef;

	if (blockRef.node.tag === "w:sectPr") {
		return commitSectionDelete(view, blockRef, opts);
	}
	return commitBlockDelete(view, blockRef, opts);
}

async function parseAndValidateOptions(
	args: string[],
): Promise<ValidatedOptions | number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: OPTION_SPEC,
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

	const filePath = parsed.positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE argument", HELP);

	const locator = parsed.values.at as string | undefined;
	if (!locator) return fail("USAGE", "Missing --at LOCATOR", HELP);

	return {
		filePath,
		locator,
		authorFlag: parsed.values.author as string | undefined,
		outputPath: parsed.values.output as string | undefined,
		dryRun: Boolean(parsed.values["dry-run"]),
	};
}

const OPTION_SPEC = {
	at: { type: "string" },
	author: { type: "string" },
	output: { type: "string", short: "o" },
	"dry-run": { type: "boolean" },
	verbose: { type: "boolean", short: "v" },
	help: { type: "boolean", short: "h" },
} as const;

type ValidatedOptions = {
	filePath: string;
	locator: string;
	authorFlag?: string;
	outputPath?: string;
	dryRun: boolean;
};

async function commitSectionDelete(
	view: DocView,
	blockRef: BlockReference,
	opts: ValidatedOptions,
): Promise<number> {
	const bodyChildren = findBodyChildren(view);
	if (bodyChildren && isTrailingSectPr(bodyChildren, blockRef.parent)) {
		return fail(
			"USAGE",
			"Cannot delete the trailing section break (mandatory in OOXML)",
			"Use `docx edit --at sN --columns 1` to reset its properties instead.",
		);
	}

	if (opts.dryRun) return respondDryRun(opts);

	const trackingOn = isTrackChangesEnabled(view);
	const owningParagraph = trackingOn
		? findContainingParagraph(view.documentTree, blockRef.node)
		: null;
	const anchorRun =
		owningParagraph?.children.find((child) => child.tag === "w:r") ?? null;

	removeInlineSectPr(blockRef.node, blockRef.parent);

	if (trackingOn && owningParagraph && anchorRun) {
		emitAuditComment(
			view,
			{ kind: "run", paragraph: owningParagraph, run: anchorRun },
			{
				body: `[docx-cli] section break removed (${opts.locator})`,
				author: resolveAuthor(opts.authorFlag),
				date: resolveDate(),
			},
		);
	}

	await saveDocView(view, opts.outputPath);
	return emitDeleteAck(opts);
}

async function commitBlockDelete(
	view: DocView,
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

	if (isTrackChangesEnabled(view)) {
		if (blockRef.node.tag !== "w:p") {
			return fail(
				"TRACKED_CHANGE_CONFLICT",
				"Tracked deletion of non-paragraph blocks (e.g., tables) is not supported",
				"Use `docx track-changes off` first, or delete table contents row-by-row.",
			);
		}
		applyTrackedDeletion(view, blockRef.node, opts.authorFlag);
	} else {
		blockRef.parent.splice(targetIndex, 1);
	}

	await saveDocView(view, opts.outputPath);
	return emitDeleteAck(opts);
}

async function respondDryRun(opts: ValidatedOptions): Promise<number> {
	await respond({
		ok: true,
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

function findBodyChildren(view: DocView): XmlNode[] | null {
	const root = XmlNode.findRoot(view.documentTree, "w:document");
	if (!root) return null;
	const body = root.findChild("w:body");
	if (!body) return null;
	return body.children;
}

function applyTrackedDeletion(
	view: DocView,
	paragraph: XmlNode,
	authorFlag: string | undefined,
): void {
	const allocator = createRevisionAllocator(view);
	const baseMeta = { author: resolveAuthor(authorFlag), date: resolveDate() };
	const mintMeta = (): TrackedMeta => ({
		...baseMeta,
		revisionId: allocator.next(),
	});

	// Wrap each contiguous run of <w:r> children in <w:del> with <w:t> -> <w:delText>.
	const newChildren: XmlNode[] = [];
	let runBuffer: XmlNode[] = [];
	const flush = (): void => {
		if (runBuffer.length === 0) return;
		const converted = runBuffer.map((run) => convertTextToDelText(run));
		newChildren.push(<Del meta={mintMeta()}>{converted}</Del>);
		runBuffer = [];
	};
	for (const child of paragraph.children) {
		if (child.tag === "w:r") {
			runBuffer.push(child);
			continue;
		}
		flush();
		newChildren.push(child);
	}
	flush();
	paragraph.children = newChildren;

	// Mark the paragraph mark itself as deleted so accepting changes also removes
	// the paragraph break, leaving no orphan empty paragraph.
	markParagraphMarkAs(paragraph, "del", mintMeta());
}
