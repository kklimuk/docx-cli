import type { Block, Paragraph } from "./types";

export function paragraphText(paragraph: Paragraph): string {
	let out = "";
	for (const run of paragraph.runs) {
		if (run.type === "text") out += run.text;
	}
	return out;
}

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
