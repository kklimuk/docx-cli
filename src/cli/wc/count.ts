import {
	type Block,
	type Doc,
	type Paragraph,
	paragraphText,
	paragraphTextAccepted,
	paragraphTextBaseline,
} from "@core";

export type CountView = "current" | "accepted" | "baseline";
export type CountOptions = { view?: CountView };

function textFor(options: CountOptions): (paragraph: Paragraph) => string {
	if (options.view === "accepted") return paragraphTextAccepted;
	if (options.view === "baseline") return paragraphTextBaseline;
	return paragraphText;
}

export function countWords(text: string): number {
	const matches = text.match(/\S+/g);
	return matches?.length ?? 0;
}

export function countWordsInDoc(doc: Doc, options: CountOptions = {}): number {
	return countWordsInBlocks(doc.blocks, options);
}

export function countWordsInBlocks(
	blocks: Block[],
	options: CountOptions = {},
): number {
	const text = textFor(options);
	let total = 0;
	for (const block of blocks) {
		if (block.type === "paragraph") {
			total += countWords(text(block));
		} else if (block.type === "table") {
			for (const row of block.rows) {
				for (const cell of row.cells) {
					total += countWordsInBlocks(cell.blocks, options);
				}
			}
		}
	}
	return total;
}

/** Word-count the half-open paragraph slice [start, end). Offsets are over
 * whichever view is selected by `options.view` — default ("current") counts
 * everything on disk, "accepted" skips del runs, "baseline" skips ins runs. */
export function countWordsInParagraphSpan(
	paragraph: Paragraph,
	start: number,
	end: number,
	options: CountOptions = {},
): number {
	const text = textFor(options)(paragraph);
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
	options: CountOptions = {},
): number {
	const startIndex = paragraphsInOrder.findIndex((p) => p.id === startBlockId);
	const endIndex = paragraphsInOrder.findIndex((p) => p.id === endBlockId);
	if (startIndex === -1 || endIndex === -1) return 0;
	if (endIndex < startIndex) return 0;

	const text = textFor(options);

	if (startIndex === endIndex) {
		const paragraph = paragraphsInOrder[startIndex];
		if (!paragraph) return 0;
		return countWordsInParagraphSpan(
			paragraph,
			startOffset,
			endOffset,
			options,
		);
	}

	let total = 0;
	const first = paragraphsInOrder[startIndex];
	if (first) {
		total += countWordsInParagraphSpan(
			first,
			startOffset,
			text(first).length,
			options,
		);
	}
	for (let index = startIndex + 1; index < endIndex; index++) {
		const middle = paragraphsInOrder[index];
		if (middle) total += countWords(text(middle));
	}
	const last = paragraphsInOrder[endIndex];
	if (last) total += countWordsInParagraphSpan(last, 0, endOffset, options);
	return total;
}
