import {
	type Block,
	findBlockById,
	flattenParagraphs,
	type Locator,
	LocatorParseError,
	paragraphText,
	parseLocator,
} from "@core";
import { parseArgs } from "util";
import { EXIT, fail, openOrFail, respond, writeStdout } from "../respond";
import {
	countWords,
	countWordsInBlocks,
	countWordsInParagraphSpan,
	countWordsInRange,
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

Options:
  -h, --help        show this help

Counting is whitespace-segmented (\\S+) over the joined paragraph text. Hidden
content like images/breaks/tabs contributes no words. Tracked-change deletions
are counted alongside insertions because their characters are still part of the
on-disk text; accept or reject changes first if you need a delta on the
"accepted" view.

Examples:
  docx wc doc.docx
  docx wc doc.docx p3
  docx wc doc.docx p3:0-120
  docx wc doc.docx p5:10-p9:42
  docx wc doc.docx t0:r1c0
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
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

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	if (!locatorInput) {
		await respond({
			ok: true,
			operation: "wc",
			path,
			scope: "document",
			words: countWordsInBlocks(view.doc.blocks),
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

	if (locator.kind === "comment" || locator.kind === "image") {
		return fail(
			"USAGE",
			`Locator ${locatorInput} addresses a ${locator.kind}, not text`,
			"wc accepts paragraph, span, range, table, and cell locators.",
		);
	}

	const blocks = view.doc.blocks;

	if (locator.kind === "block") {
		const block = findBlockById(blocks, locator.blockId);
		if (!block) {
			return fail("BLOCK_NOT_FOUND", `Block not found: ${locator.blockId}`);
		}
		const words =
			block.type === "paragraph"
				? countWords(paragraphText(block))
				: block.type === "table"
					? countWordsInBlocks([block])
					: 0;
		await respond({
			ok: true,
			operation: "wc",
			path,
			locator: locatorInput,
			scope: block.type,
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
			words: countWordsInParagraphSpan(block, locator.start, locator.end),
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
			words: countWordsInRange(
				paragraphs,
				locator.start.blockId,
				locator.start.offset,
				locator.end.blockId,
				locator.end.offset,
			),
		});
		return EXIT.OK;
	}

	if (locator.kind === "cell") {
		const cellPath = `${locator.tableId}:r${locator.row}c${locator.col}`;
		if (!locator.inner) {
			const cellBlocks = findCellBlocks(blocks, locator);
			if (!cellBlocks) {
				return fail("BLOCK_NOT_FOUND", `Cell not found: ${cellPath}`);
			}
			await respond({
				ok: true,
				operation: "wc",
				path,
				locator: locatorInput,
				scope: "cell",
				words: countWordsInBlocks(cellBlocks),
			});
			return EXIT.OK;
		}
		// inner addresses a paragraph or paragraph span inside the cell;
		// the AST gives those a flattened id like "t0:r1c0:p0", so resolve via that.
		const innerBlockId =
			locator.inner.kind === "block"
				? locator.inner.blockId
				: locator.inner.kind === "blockSpan"
					? locator.inner.blockId
					: null;
		if (!innerBlockId) {
			return fail(
				"USAGE",
				`Unsupported inner locator for cell: ${locatorInput}`,
				"Inside a cell, wc accepts pK or pK:S-E only.",
			);
		}
		const composedId = `${cellPath}:${innerBlockId}`;
		const block = findBlockById(blocks, composedId);
		if (!block || block.type !== "paragraph") {
			return fail("BLOCK_NOT_FOUND", `Paragraph not found: ${composedId}`);
		}
		const words =
			locator.inner.kind === "blockSpan"
				? countWordsInParagraphSpan(
						block,
						locator.inner.start,
						locator.inner.end,
					)
				: countWords(paragraphText(block));
		await respond({
			ok: true,
			operation: "wc",
			path,
			locator: locatorInput,
			scope: locator.inner.kind === "blockSpan" ? "paragraphSpan" : "paragraph",
			words,
		});
		return EXIT.OK;
	}

	return fail("USAGE", `Unsupported locator: ${locatorInput}`);
}

function findCellBlocks(
	blocks: Block[],
	locator: Extract<Locator, { kind: "cell" }>,
): Block[] | null {
	const block = findBlockById(blocks, locator.tableId);
	if (!block || block.type !== "table") return null;
	const row = block.rows[locator.row];
	if (!row) return null;
	const cell = row.cells[locator.col];
	if (!cell) return null;
	return cell.blocks;
}
