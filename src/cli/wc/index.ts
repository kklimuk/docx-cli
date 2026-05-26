import {
	type Block,
	findBlockById,
	flattenParagraphs,
	type Locator,
	LocatorParseError,
	type Paragraph,
	paragraphText,
	paragraphTextAccepted,
	paragraphTextBaseline,
	parseLocator,
} from "@core";
import { parseArgs } from "util";
import { EXIT, fail, openOrFail, respond, writeStdout } from "../respond";
import {
	type CountView,
	countSectionsInBlocks,
	countWords,
	countWordsInBlockRange,
	countWordsInBlocks,
	countWordsInParagraphSpan,
	countWordsInRange,
	countWordsInSection,
} from "./count";

const HELP = `docx wc — count words in a document or a locator-addressed slice

Usage:
  docx wc FILE [LOCATOR] [options]

Locators (optional; default: whole document):
  pN              whole paragraph N
  pN:S-E          chars S..E within paragraph N
  pN:S-pM:E       cross-paragraph range
  tN              whole table
  tN:rRcC         whole cell
  tN:rRcC:pK      paragraph K inside that cell
  tN:rRcC:pK:S-E  span within a cell paragraph
  tN:rRcC:tK[…]   nested table inside a cell — chain further (\`:rRcC\`,
                  \`:pK\`, \`:pK:S-E\`) to address rows/cells/paragraphs/spans
                  at any depth
  sN              section N — every paragraph and table from the prior section
                  boundary up to and including the paragraph that holds sN's
                  inline sectPr (or to end of body for the trailing section)

Options:
  --accepted        Default. Count the accepted view: skip subtractive
                    wrappers (<w:del>, <w:moveFrom>); keep additive
                    wrappers (<w:ins>, <w:moveTo>) as plain text. Mirrors
                    "docx read" / "docx find" defaults.
  --baseline        Count the baseline view: skip additive wrappers
                    (<w:ins>, <w:moveTo>); keep subtractive wrappers
                    (<w:del>, <w:moveFrom>) as plain text — i.e., the doc as
                    it was before any tracked changes were made.
  --current         Count the raw concatenation: every tracked-change
                    wrapper's text counts (everything on disk).
  -h, --help        show this help

Counting is whitespace-segmented (\\S+) over the joined paragraph text. Hidden
content like images/breaks/tabs contributes no words. The three view flags
are mutually exclusive; default (no flag) is --accepted, matching what a
reader would see if every tracked change were accepted.

Examples:
  docx wc doc.docx
  docx wc doc.docx p3
  docx wc doc.docx p3:0-120
  docx wc doc.docx p5:10-p9:42
  docx wc doc.docx t0:r1c0
  docx wc doc.docx s2
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				accepted: { type: "boolean" },
				baseline: { type: "boolean" },
				current: { type: "boolean" },
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

	const path = parsed.positionals[0];
	const locatorInput = parsed.positionals[1];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const wantAccepted = Boolean(parsed.values.accepted);
	const wantBaseline = Boolean(parsed.values.baseline);
	const wantCurrent = Boolean(parsed.values.current);
	const viewFlagCount =
		(wantAccepted ? 1 : 0) + (wantBaseline ? 1 : 0) + (wantCurrent ? 1 : 0);
	if (viewFlagCount > 1) {
		return fail(
			"USAGE",
			"--accepted, --baseline, and --current are mutually exclusive",
			HELP,
		);
	}
	const view: CountView = wantCurrent
		? "current"
		: wantBaseline
			? "baseline"
			: "accepted";
	const pickText = paragraphTextFor(view);

	const docView = await openOrFail(path);
	if (typeof docView === "number") return docView;

	if (!locatorInput) {
		await respond({
			ok: true,
			operation: "wc",
			path,
			scope: "document",
			view,
			words: countWordsInBlocks(docView.doc.blocks, { view }),
			sections: countSectionsInBlocks(docView.doc.blocks),
		});
		return EXIT.OK;
	}

	let locator: Locator;
	try {
		locator = parseLocator(locatorInput);
	} catch (error) {
		if (error instanceof LocatorParseError) {
			return fail("INVALID_LOCATOR", error.message);
		}
		throw error;
	}

	if (
		locator.kind === "comment" ||
		locator.kind === "image" ||
		locator.kind === "hyperlink" ||
		locator.kind === "trackedChange" ||
		locator.kind === "equation"
	) {
		return fail(
			"USAGE",
			`Locator ${locatorInput} addresses a ${locator.kind}, not text`,
			"wc accepts paragraph, span, range, table, and cell locators.",
		);
	}

	const blocks = docView.doc.blocks;

	if (locator.kind === "block") {
		const block = findBlockById(blocks, locator.blockId);
		if (!block) {
			return fail("BLOCK_NOT_FOUND", `Block not found: ${locator.blockId}`);
		}
		if (block.type === "sectionBreak") {
			const sectionWords = countWordsInSection(blocks, locator.blockId, {
				view,
			});
			if (sectionWords === null) {
				return fail("BLOCK_NOT_FOUND", `Section not found: ${locator.blockId}`);
			}
			await respond({
				ok: true,
				operation: "wc",
				path,
				locator: locatorInput,
				scope: "section",
				view,
				words: sectionWords,
			});
			return EXIT.OK;
		}
		const words =
			block.type === "paragraph"
				? countWords(pickText(block))
				: block.type === "table"
					? countWordsInBlocks([block], { view })
					: 0;
		await respond({
			ok: true,
			operation: "wc",
			path,
			locator: locatorInput,
			scope: block.type,
			view,
			words,
		});
		return EXIT.OK;
	}

	if (locator.kind === "blockSpan") {
		const block = findBlockById(blocks, locator.blockId);
		if (!block || block.type !== "paragraph") {
			return fail("BLOCK_NOT_FOUND", `Paragraph not found: ${locator.blockId}`);
		}
		await respond({
			ok: true,
			operation: "wc",
			path,
			locator: locatorInput,
			scope: "paragraphSpan",
			view,
			words: countWordsInParagraphSpan(block, locator.start, locator.end, {
				view,
			}),
		});
		return EXIT.OK;
	}

	if (locator.kind === "blockRange") {
		const paragraphs = flattenParagraphs(blocks);
		const startExists = paragraphs.some((p) => p.id === locator.startBlockId);
		const endExists = paragraphs.some((p) => p.id === locator.endBlockId);
		if (!startExists) {
			return fail(
				"BLOCK_NOT_FOUND",
				`Paragraph not found: ${locator.startBlockId}`,
			);
		}
		if (!endExists) {
			return fail(
				"BLOCK_NOT_FOUND",
				`Paragraph not found: ${locator.endBlockId}`,
			);
		}
		await respond({
			ok: true,
			operation: "wc",
			path,
			locator: locatorInput,
			scope: "blockRange",
			view,
			words: countWordsInBlockRange(
				paragraphs,
				locator.startBlockId,
				locator.endBlockId,
				{ view },
			),
		});
		return EXIT.OK;
	}

	if (locator.kind === "range") {
		const paragraphs = flattenParagraphs(blocks);
		const startExists = paragraphs.some((p) => p.id === locator.start.blockId);
		const endExists = paragraphs.some((p) => p.id === locator.end.blockId);
		if (!startExists) {
			return fail(
				"BLOCK_NOT_FOUND",
				`Paragraph not found: ${locator.start.blockId}`,
			);
		}
		if (!endExists) {
			return fail(
				"BLOCK_NOT_FOUND",
				`Paragraph not found: ${locator.end.blockId}`,
			);
		}
		await respond({
			ok: true,
			operation: "wc",
			path,
			locator: locatorInput,
			scope: "range",
			view,
			words: countWordsInRange(
				paragraphs,
				locator.start.blockId,
				locator.start.offset,
				locator.end.blockId,
				locator.end.offset,
				{ view },
			),
		});
		return EXIT.OK;
	}

	if (locator.kind === "cell") {
		const resolved = resolveCellChain(blocks, locator);
		if (!resolved) {
			return fail("BLOCK_NOT_FOUND", `Not found: ${locatorInput}`);
		}
		if (resolved.kind === "wholeCell") {
			await respond({
				ok: true,
				operation: "wc",
				path,
				locator: locatorInput,
				scope: "cell",
				view,
				words: countWordsInBlocks(resolved.blocks, { view }),
			});
			return EXIT.OK;
		}
		if (resolved.kind === "wholeTable") {
			await respond({
				ok: true,
				operation: "wc",
				path,
				locator: locatorInput,
				scope: "table",
				view,
				words: countWordsInBlocks([resolved.block], { view }),
			});
			return EXIT.OK;
		}
		const words = resolved.span
			? countWordsInParagraphSpan(
					resolved.paragraph,
					resolved.span.start,
					resolved.span.end,
					{ view },
				)
			: countWords(pickText(resolved.paragraph));
		await respond({
			ok: true,
			operation: "wc",
			path,
			locator: locatorInput,
			scope: resolved.span ? "paragraphSpan" : "paragraph",
			view,
			words,
		});
		return EXIT.OK;
	}

	return fail("USAGE", `Unsupported locator: ${locatorInput}`);
}

/** Walk a (possibly nested) cell locator chain against `blocks`, the block list
 * at the current depth. At each level, the locator's `tableId` / `blockId`
 * (e.g. "t0", "p3") is a *local* index over blocks of that type at this depth,
 * not a globally-qualified id — that matches what `parseLocator` produces for
 * the inner segments. Returns the leaf the locator addresses: a whole cell, a
 * whole (possibly nested) table, or a single paragraph with optional span. */
type ResolvedCellTarget =
	| { kind: "wholeCell"; blocks: Block[] }
	| { kind: "wholeTable"; block: Block }
	| {
			kind: "paragraph";
			paragraph: Paragraph;
			span?: { start: number; end: number };
	  };

function resolveCellChain(
	blocks: Block[],
	locator: Locator,
): ResolvedCellTarget | null {
	if (locator.kind === "cell") {
		const table = findNthBlockOfKind(blocks, "table", locator.tableId);
		if (!table || table.type !== "table") return null;
		const row = table.rows[locator.row];
		if (!row) return null;
		const cell = row.cells[locator.col];
		if (!cell) return null;
		if (!locator.inner) return { kind: "wholeCell", blocks: cell.blocks };
		return resolveCellChain(cell.blocks, locator.inner);
	}
	if (locator.kind === "block") {
		if (locator.blockId.startsWith("p")) {
			const paragraph = findNthBlockOfKind(
				blocks,
				"paragraph",
				locator.blockId,
			);
			if (!paragraph || paragraph.type !== "paragraph") return null;
			return { kind: "paragraph", paragraph };
		}
		if (locator.blockId.startsWith("t")) {
			const table = findNthBlockOfKind(blocks, "table", locator.blockId);
			if (!table) return null;
			return { kind: "wholeTable", block: table };
		}
		return null;
	}
	if (locator.kind === "blockSpan") {
		const paragraph = findNthBlockOfKind(blocks, "paragraph", locator.blockId);
		if (!paragraph || paragraph.type !== "paragraph") return null;
		return {
			kind: "paragraph",
			paragraph,
			span: { start: locator.start, end: locator.end },
		};
	}
	return null;
}

/** Find the Nth block of `kind` in `blocks`, where the id's numeric suffix is
 * the 0-based position among blocks of that type at this depth — e.g. "t0" =
 * first table, "p2" = third paragraph. Returns null if out of range. */
function findNthBlockOfKind(
	blocks: Block[],
	kind: "paragraph" | "table",
	idWithPrefix: string,
): Block | null {
	const index = Number.parseInt(idWithPrefix.slice(1), 10);
	if (!Number.isInteger(index) || index < 0) return null;
	let seen = 0;
	for (const block of blocks) {
		if (block.type !== kind) continue;
		if (seen === index) return block;
		seen++;
	}
	return null;
}

function paragraphTextFor(view: CountView): (p: Paragraph) => string {
	if (view === "accepted") return paragraphTextAccepted;
	if (view === "baseline") return paragraphTextBaseline;
	return paragraphText;
}
