import { type Block, type Doc, type Paragraph, paragraphText } from "@core";

export type OutlineEntry = {
	id: string;
	locator: string;
	level: number;
	style: string;
	text: string;
	children: OutlineEntry[];
};

export type OutlineOptions = {
	stylePrefix?: string;
};

export function buildOutline(
	doc: Doc,
	options: OutlineOptions = {},
): OutlineEntry[] {
	const stylePrefix = options.stylePrefix ?? "Heading";
	const root: OutlineEntry = {
		id: "",
		locator: "",
		level: 0,
		style: "",
		text: "",
		children: [],
	};
	const stack: OutlineEntry[] = [root];

	for (const paragraph of headingParagraphs(doc.blocks, stylePrefix)) {
		const level = headingLevel(paragraph.style, stylePrefix);
		if (level === null) continue;
		while ((stack[stack.length - 1]?.level ?? 0) >= level) stack.pop();
		const parent = stack[stack.length - 1] ?? root;
		const entry: OutlineEntry = {
			id: paragraph.id,
			locator: paragraph.id,
			level,
			style: paragraph.style ?? "",
			text: paragraphText(paragraph),
			children: [],
		};
		parent.children.push(entry);
		stack.push(entry);
	}

	return root.children;
}

function headingLevel(
	style: string | undefined,
	stylePrefix: string,
): number | null {
	if (!style) return null;
	if (!style.startsWith(stylePrefix)) return null;
	const remainder = style.slice(stylePrefix.length).trim();
	if (remainder === "") return 1;
	const parsed = Number(remainder);
	if (!Number.isInteger(parsed) || parsed < 1) return null;
	return parsed;
}

function* headingParagraphs(
	blocks: Block[],
	stylePrefix: string,
): Generator<Paragraph> {
	for (const block of blocks) {
		if (block.type !== "paragraph") continue;
		if (headingLevel(block.style, stylePrefix) === null) continue;
		yield block;
	}
}
