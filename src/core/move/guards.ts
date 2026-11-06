import type { XmlNode } from "../parser";
import type { MoveDestination } from "./destination";
import { MoveError } from "./move-error";
import type { MoveSource } from "./source";

/** Refuse a source whose move would change more than its position: a
 *  paragraph carrying a section break would drag the section boundary along,
 *  and a range marker whose other end sits outside the source would be split
 *  or inverted by the move. */
export function assertMovableSource(source: MoveSource): void {
	const carriesSectPr = (node: XmlNode) =>
		node.tag === "w:sectPr" || node.findDescendant("w:sectPr") !== undefined;
	if (source.nodes.some(carriesSectPr)) {
		throw new MoveError(
			"UNSUPPORTED",
			`${source.locator} ends a section (it carries a section break), so moving it would move the section boundary`,
			"Move the blocks around it instead, or change the layout with `docx sections`.",
		);
	}
	const straddler = straddlingRange(source.nodes);
	if (straddler) {
		throw new MoveError(
			"UNSUPPORTED",
			`${straddler.what} starts on one side of the edge of ${source.locator} and ends on the other — moving it would split that range`,
			`Widen --at to a range (pN-pM) that holds both ends, or ${straddler.cure}`,
		);
	}
}

/** Moves stay inside one story: the body (cells included) or one text box.
 *  A text box's story is anchored by a body paragraph and mirrored into its
 *  Fallback twin, so carrying blocks across that boundary is out of scope. */
export function assertSameStory(
	source: MoveSource,
	destination: MoveDestination,
): void {
	if (source.story === destination.story) return;
	throw new MoveError(
		"UNSUPPORTED",
		`Can't move ${source.locator} (${describeStory(source.story)}) to ${destination.locator} (${describeStory(destination.story)})`,
		"move works within the body (table cells included) or within one text box. Across that boundary, insert the content at the destination and delete the original.",
	);
}

export function assertDestinationOutsideSource(
	source: MoveSource,
	destination: MoveDestination,
): void {
	if (!destinationInside(destination, source)) return;
	throw new MoveError(
		"USAGE",
		`${destination.locator} is part of what's being moved (${source.locator}), so it can't be the destination`,
		"Anchor on a block outside the blocks being moved.",
	);
}

/** A move that lands the source where it already is changes nothing — an
 *  error, never a silent success. Judged on block siblings (paragraphs,
 *  tables, content controls), the order `read` shows. */
export function assertNotNoOp(
	source: MoveSource,
	destination: MoveDestination,
): void {
	const sameContainer =
		destination.kind === "cell"
			? destination.cell.parent === source.parent
			: destination.parent === source.parent;
	if (!sameContainer) return;
	const blocks = source.parent.filter(isBlock);
	const sourceBlocks = source.nodes.filter(isBlock);
	const first = blocks.indexOf(sourceBlocks[0] as XmlNode);
	const last = blocks.indexOf(sourceBlocks[sourceBlocks.length - 1] as XmlNode);
	const lands =
		destination.kind === "cell"
			? destination.side === "before"
				? first === 0
				: last === blocks.length - 1
			: destination.side === "after"
				? blocks.indexOf(destination.anchor) === first - 1
				: blocks.indexOf(destination.anchor) === last + 1;
	if (!lands) return;
	throw new MoveError(
		"USAGE",
		`${source.locator} is already ${destination.side} ${destination.locator} — nothing to move`,
		"Pick a destination elsewhere; `docx read` shows the current order.",
	);
}

/** Batch: two entries may not move the same block, or a block inside another
 *  entry's block. */
export function assertSourcesDisjoint(
	earlier: MoveSource,
	later: MoveSource,
): void {
	const overlaps = earlier.nodes.some((outer) =>
		later.nodes.some(
			(inner) => containsNode(outer, inner) || containsNode(inner, outer),
		),
	);
	if (!overlaps) return;
	throw new MoveError(
		"USAGE",
		`${later.locator} overlaps ${earlier.locator}, which an earlier entry already moves`,
		"Each entry must move distinct blocks; merge the two into one entry.",
	);
}

/** Whether the destination's anchor (a block, or a bare cell's `<w:tc>`) is
 *  one of the source nodes or sits inside one. */
export function destinationInside(
	destination: MoveDestination,
	source: MoveSource,
): boolean {
	const anchor =
		destination.kind === "cell" ? destination.cell.node : destination.anchor;
	return source.nodes.some((node) => containsNode(node, anchor));
}

function containsNode(root: XmlNode, target: XmlNode): boolean {
	if (root === target) return true;
	return root.children.some((child) => containsNode(child, target));
}

const BLOCK_TAGS = new Set(["w:p", "w:tbl", "w:sdt"]);

function isBlock(node: XmlNode): boolean {
	return BLOCK_TAGS.has(node.tag);
}

/** Paired range markers keyed by `w:id`: a pair with exactly one end inside
 *  the source straddles its edge. */
const RANGE_PAIRS = [
	{
		start: "w:commentRangeStart",
		end: "w:commentRangeEnd",
		what: (id: string) => `Comment c${id}`,
		cure: "re-anchor the comment first (`docx comments delete` + `comments add`).",
	},
	{
		start: "w:bookmarkStart",
		end: "w:bookmarkEnd",
		what: (id: string) => `Bookmark ${id}`,
		cure: "move a range that keeps the bookmark whole.",
	},
	{
		start: "w:moveFromRangeStart",
		end: "w:moveFromRangeEnd",
		what: (id: string) => `Tracked move range ${id}`,
		cure: "resolve the pending move first (`docx track-changes list`).",
	},
	{
		start: "w:moveToRangeStart",
		end: "w:moveToRangeEnd",
		what: (id: string) => `Tracked move range ${id}`,
		cure: "resolve the pending move first (`docx track-changes list`).",
	},
	{
		start: "w:permStart",
		end: "w:permEnd",
		what: (id: string) => `Editing-permission range ${id}`,
		cure: "move a range that keeps the permission range whole.",
	},
] as const;

function straddlingRange(
	nodes: XmlNode[],
): { what: string; cure: string } | null {
	for (const pair of RANGE_PAIRS) {
		const starts = new Map<string, XmlNode>();
		const ends = new Map<string, XmlNode>();
		for (const node of nodes) {
			collectIds(node, pair.start, starts);
			collectIds(node, pair.end, ends);
		}
		for (const [id, marker] of starts) {
			if (ends.has(id)) continue;
			return { what: label(pair, id, marker), cure: pair.cure };
		}
		for (const id of ends.keys()) {
			if (starts.has(id)) continue;
			return { what: pair.what(id), cure: pair.cure };
		}
	}
	return null;
}

function label(
	pair: (typeof RANGE_PAIRS)[number],
	id: string,
	marker: XmlNode,
): string {
	const name = marker.getAttribute("w:name");
	return name && pair.start === "w:bookmarkStart"
		? `Bookmark "${name}"`
		: pair.what(id);
}

function collectIds(
	node: XmlNode,
	tag: string,
	into: Map<string, XmlNode>,
): void {
	const id = node.tag === tag ? node.getAttribute("w:id") : undefined;
	if (id !== undefined) into.set(id, node);
	for (const child of node.children) collectIds(child, tag, into);
}

function describeStory(story: string): string {
	return story === "body" ? "the body" : `text box ${story}`;
}
