import type { Block, Paragraph } from "./types";

export function paragraphText(paragraph: Paragraph): string {
	let out = "";
	for (const run of paragraph.runs) {
		if (run.type === "text") out += run.text;
	}
	return out;
}

/** Concatenate text as it would read in the accepted view: skip runs inside
 * a tracked deletion (`<w:del>`), keep tracked insertions. */
export function paragraphTextAccepted(paragraph: Paragraph): string {
	let out = "";
	for (const run of paragraph.runs) {
		if (run.type !== "text") continue;
		if (run.trackedChange?.kind === "del") continue;
		out += run.text;
	}
	return out;
}

/** Concatenate text as it would read in the baseline (pre-change) view: skip
 * runs inside a tracked insertion (`<w:ins>`), keep tracked deletions. */
export function paragraphTextBaseline(paragraph: Paragraph): string {
	let out = "";
	for (const run of paragraph.runs) {
		if (run.type !== "text") continue;
		if (run.trackedChange?.kind === "ins") continue;
		out += run.text;
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
