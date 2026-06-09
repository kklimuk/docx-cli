import {
	type Block,
	describeForms,
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
import { resolveView } from "../parse-helpers";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";
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

Locator (optional positional; default: whole document):
${describeForms([
	"paragraph",
	"blockRange",
	"span",
	"crossSpan",
	"table",
	"cell",
	"cellParagraph",
	"cellSpan",
	"section",
])}
  cellParagraph chains to any nesting depth (a table inside a cell:
  tN:rRcC:tN:rR2cC2:pK …). Notation: uppercase letters are numeric indices;
  offsets 0-based, end-exclusive. Section sN counts every paragraph and table
  from the prior section boundary through the paragraph holding sN's inline
  sectPr (or to end of body for the trailing section). Table row/column
  (tN:rR, tN:cC) and cell-range forms are NOT accepted — use tN or tN:rRcC.
  See \`docx info locators\`.

View flags (mutually exclusive; default --accepted):
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

Options:
  --json            Emit JSON instead of the bare count
  -h, --help        show this help

Counting is whitespace-segmented (\\S+) over the joined paragraph text. Hidden
content like images/breaks/tabs contributes no words.

Output:
  Default: the bare word count (a single integer). For the whole document a
  second tab-separated column gives the section count, like real \`wc\`.
  --json: { words, scope, view, sections? } (no envelope). Errors print
  {code, error, hint?} with a nonzero exit.

Examples:
  docx wc doc.docx
  docx wc doc.docx p3
  docx wc doc.docx p2-p5
  docx wc doc.docx p3:0-120
  docx wc doc.docx p5:10-p9:42
  docx wc doc.docx t0:r1c0
  docx wc doc.docx s2
  docx wc doc.docx --json | jq .words
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			accepted: { type: "boolean" },
			baseline: { type: "boolean" },
			current: { type: "boolean" },
			json: { type: "boolean" },
			help: { type: "boolean", short: "h" },
		},
		HELP,
	);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	const path = parsed.positionals[0];
	const locatorInput = parsed.positionals[1];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const view = resolveView(parsed.values);
	if (!view) {
		return fail(
			"USAGE",
			"--accepted, --baseline, and --current are mutually exclusive",
			HELP,
		);
	}
	const json = Boolean(parsed.values.json);
	const pickText = paragraphTextFor(view);

	const docView = await openOrFail(path);
	if (typeof docView === "number") return docView;

	if (!locatorInput) {
		return emitWc(json, countWordsInBlocks(docView.body.blocks, { view }), {
			scope: "document",
			view,
			sections: countSectionsInBlocks(docView.body.blocks),
		});
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
		locator.kind === "equation" ||
		locator.kind === "footnote" ||
		locator.kind === "endnote"
	) {
		return fail(
			"USAGE",
			`Locator ${locatorInput} addresses a ${locator.kind}, not text`,
			"wc accepts paragraph, span, range, section, table, and cell locators.",
		);
	}

	const blocks = docView.body.blocks;

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
			return emitWc(json, sectionWords, { scope: "section", view });
		}
		const words =
			block.type === "paragraph"
				? countWords(pickText(block))
				: block.type === "table"
					? countWordsInBlocks([block], { view })
					: 0;
		return emitWc(json, words, { scope: block.type, view });
	}

	if (locator.kind === "blockSpan") {
		const block = findBlockById(blocks, locator.blockId);
		if (!block || block.type !== "paragraph") {
			return fail("BLOCK_NOT_FOUND", `Paragraph not found: ${locator.blockId}`);
		}
		return emitWc(
			json,
			countWordsInParagraphSpan(block, locator.start, locator.end, { view }),
			{ scope: "paragraphSpan", view },
		);
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
		return emitWc(
			json,
			countWordsInBlockRange(
				paragraphs,
				locator.startBlockId,
				locator.endBlockId,
				{ view },
			),
			{ scope: "blockRange", view },
		);
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
		return emitWc(
			json,
			countWordsInRange(
				paragraphs,
				locator.start.blockId,
				locator.start.offset,
				locator.end.blockId,
				locator.end.offset,
				{ view },
			),
			{ scope: "range", view },
		);
	}

	if (locator.kind === "cell") {
		const resolved = resolveCellChain(blocks, locator);
		if (!resolved) {
			return fail("BLOCK_NOT_FOUND", `Not found: ${locatorInput}`);
		}
		if (resolved.kind === "wholeCell") {
			return emitWc(json, countWordsInBlocks(resolved.blocks, { view }), {
				scope: "cell",
				view,
			});
		}
		if (resolved.kind === "wholeTable") {
			return emitWc(json, countWordsInBlocks([resolved.block], { view }), {
				scope: "table",
				view,
			});
		}
		const words = resolved.span
			? countWordsInParagraphSpan(
					resolved.paragraph,
					resolved.span.start,
					resolved.span.end,
					{ view },
				)
			: countWords(pickText(resolved.paragraph));
		return emitWc(json, words, {
			scope: resolved.span ? "paragraphSpan" : "paragraph",
			view,
		});
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

/** wc output: bare integer by default (whole-doc adds a tab-separated section
 *  count, like real `wc`); `--json` emits the structured payload (no envelope). */
async function emitWc(
	json: boolean,
	words: number,
	meta: { scope: string; view: CountView; sections?: number },
): Promise<number> {
	if (json) {
		await respond({
			words,
			scope: meta.scope,
			view: meta.view,
			...(meta.sections !== undefined ? { sections: meta.sections } : {}),
		});
	} else {
		await writeStdout(
			meta.sections !== undefined
				? `${words}\t${meta.sections}\n`
				: `${words}\n`,
		);
	}
	return EXIT.OK;
}

function paragraphTextFor(view: CountView): (p: Paragraph) => string {
	if (view === "accepted") return paragraphTextAccepted;
	if (view === "baseline") return paragraphTextBaseline;
	return paragraphText;
}
