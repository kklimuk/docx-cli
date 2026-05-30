import { iterateBlocks } from "./document/body";
import type { Block, ImageRun, Paragraph } from "./types";

export function paragraphText(paragraph: Paragraph): string {
	let out = "";
	for (const run of paragraph.runs) {
		if (run.type === "text") out += run.text;
	}
	return out;
}

/** Concatenate text as it would read in the accepted view: skip runs inside
 * a tracked deletion (`<w:del>`) or a tracked-move source (`<w:moveFrom>`);
 * keep insertions and move destinations. */
export function paragraphTextAccepted(paragraph: Paragraph): string {
	let out = "";
	for (const run of paragraph.runs) {
		if (run.type !== "text") continue;
		const kind = run.trackedChange?.kind;
		if (kind === "del" || kind === "moveFrom") continue;
		out += run.text;
	}
	return out;
}

/** Concatenate text as it would read in the baseline (pre-change) view: skip
 * runs inside a tracked insertion (`<w:ins>`) or a tracked-move destination
 * (`<w:moveTo>`); keep deletions and move sources. */
export function paragraphTextBaseline(paragraph: Paragraph): string {
	let out = "";
	for (const run of paragraph.runs) {
		if (run.type !== "text") continue;
		const kind = run.trackedChange?.kind;
		if (kind === "ins" || kind === "moveTo") continue;
		out += run.text;
	}
	return out;
}

export function flattenParagraphs(blocks: Block[]): Paragraph[] {
	const out: Paragraph[] = [];
	for (const block of iterateBlocks(blocks)) {
		if (block.type === "paragraph") out.push(block);
	}
	return out;
}

/** Every image run in the document, in reading order (descending into table
 * cells) — what `images list`/`extract` enumerate. */
export function flattenImageRuns(blocks: Block[]): ImageRun[] {
	return flattenParagraphs(blocks)
		.flatMap((paragraph) => paragraph.runs)
		.filter((run): run is ImageRun => run.type === "image");
}

export function findBlockById(blocks: Block[], blockId: string): Block | null {
	for (const block of iterateBlocks(blocks)) {
		if (block.id === blockId) return block;
	}
	return null;
}
