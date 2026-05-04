import type { Block, Doc, Paragraph } from "@core";

export function countWords(text: string): number {
	const matches = text.match(/\S+/g);
	return matches?.length ?? 0;
}

export function paragraphText(paragraph: Paragraph): string {
	let out = "";
	for (const run of paragraph.runs) {
		if (run.type === "text") out += run.text;
	}
	return out;
}

export function countWordsInDoc(doc: Doc): number {
	return countWordsInBlocks(doc.blocks);
}

export function countWordsInBlocks(blocks: Block[]): number {
	let total = 0;
	for (const block of blocks) {
		if (block.type === "paragraph") {
			total += countWords(paragraphText(block));
		} else if (block.type === "table") {
			for (const row of block.rows) {
				for (const cell of row.cells) {
					total += countWordsInBlocks(cell.blocks);
				}
			}
		}
	}
	return total;
}

export function findBlockById(blocks: Block[], blockId: string): Block | null {
	for (const block of blocks) {
		if (block.id === blockId) return block;
		if (block.type === "table") {
			for (const row of block.rows) {
				for (const cell of row.cells) {
					const inner = findBlockById(cell.blocks, blockId);
					if (inner) return inner;
				}
			}
		}
	}
	return null;
}

/** Word-count the half-open paragraph slice [start, end). */
export function countWordsInParagraphSpan(
	paragraph: Paragraph,
	start: number,
	end: number,
): number {
	const text = paragraphText(paragraph);
	const clampedEnd = Math.min(Math.max(end, 0), text.length);
	const clampedStart = Math.min(Math.max(start, 0), clampedEnd);
	return countWords(text.slice(clampedStart, clampedEnd));
}

/** Word-count a cross-paragraph range: chars [startOffset..end-of-pStart], all
 * intervening paragraphs in document order, and chars [0..endOffset) of pEnd.
 * `paragraphsInOrder` is the flat list of paragraph blocks the range may span. */
export function countWordsInRange(
	paragraphsInOrder: Paragraph[],
	startBlockId: string,
	startOffset: number,
	endBlockId: string,
	endOffset: number,
): number {
	const startIndex = paragraphsInOrder.findIndex((p) => p.id === startBlockId);
	const endIndex = paragraphsInOrder.findIndex((p) => p.id === endBlockId);
	if (startIndex === -1 || endIndex === -1) return 0;
	if (endIndex < startIndex) return 0;

	if (startIndex === endIndex) {
		const paragraph = paragraphsInOrder[startIndex];
		if (!paragraph) return 0;
		return countWordsInParagraphSpan(paragraph, startOffset, endOffset);
	}

	let total = 0;
	const first = paragraphsInOrder[startIndex];
	if (first) {
		total += countWordsInParagraphSpan(
			first,
			startOffset,
			paragraphText(first).length,
		);
	}
	for (let index = startIndex + 1; index < endIndex; index++) {
		const middle = paragraphsInOrder[index];
		if (middle) total += countWords(paragraphText(middle));
	}
	const last = paragraphsInOrder[endIndex];
	if (last) total += countWordsInParagraphSpan(last, 0, endOffset);
	return total;
}

/** Flatten a doc to its paragraphs in document order, including paragraphs
 * nested in table cells. Used for cross-paragraph range counting. */
export function flattenParagraphs(blocks: Block[]): Paragraph[] {
	const out: Paragraph[] = [];
	for (const block of blocks) {
		if (block.type === "paragraph") {
			out.push(block);
			continue;
		}
		if (block.type === "table") {
			for (const row of block.rows) {
				for (const cell of row.cells) {
					out.push(...flattenParagraphs(cell.blocks));
				}
			}
		}
	}
	return out;
}
