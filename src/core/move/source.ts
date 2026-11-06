import type { Document } from "../ast/document";
import { headingLevel } from "../heading-level";
import { parseLocator } from "../locators/parse";
import { parseCellAt } from "../locators/resolve";
import type { XmlNode } from "../parser";
import { MoveError } from "./move-error";

/** The contiguous run of sibling nodes a move lifts out: `nodes` live in
 *  `parent` (a body, cell or text-box story child list) in document order.
 *  `story` is the text-box story the source lives in (`tbxN`), or `body`. */
export type MoveSource = {
	locator: string;
	parent: XmlNode[];
	nodes: XmlNode[];
	story: string;
};

/** Resolve `--at` to the live nodes it names: a paragraph or table (`pN`,
 *  `tN`, a cell paragraph `tN:rRcC:pK`, a text-box paragraph `tbxN:pK`), a
 *  sibling range `pN-pM` (every node between the endpoints, tables included),
 *  or — with `subtree` — a heading plus the blocks under it. */
export function resolveMoveSource(
	document: Document,
	locator: string,
	subtree: boolean,
): MoveSource {
	const parsed = parseLocator(locator);
	if (parsed.kind === "blockRange") {
		if (subtree) {
			throw new MoveError(
				"USAGE",
				"--subtree expands a single heading; it doesn't combine with a range",
				"Pass the heading's pN alone (`--at pN --subtree`), or drop --subtree.",
			);
		}
		const range = document.body.resolveBlockRange(
			parsed.startBlockId,
			parsed.endBlockId,
		);
		return {
			locator,
			parent: range.parent,
			nodes: range.parent.slice(range.startIndex, range.endIndex + 1),
			story: storyOf(locator),
		};
	}
	if (/:\d+-\d+$/.test(locator) || parseCellAt(locator)) {
		throw new MoveError(
			"INVALID_LOCATOR",
			`move relocates whole blocks; ${locator} isn't one`,
			"Pass a paragraph or table locator (pN, pN-pM, tN, tN:rRcC:pK, tbxN:pK) — for a cell, its paragraph (tN:rRcC:pK). To move part of a paragraph's text, edit it out and insert it at the destination.",
		);
	}
	const reference = document.body.resolveBlock(locator);
	if (reference.node.tag === "w:sectPr") {
		throw new MoveError(
			"UNSUPPORTED",
			`${locator} is a section break, which can't be moved`,
			"Section layout changes with `docx sections`; move the content around it instead.",
		);
	}
	const nodes = subtree
		? headingSubtree(reference.node, reference.parent, locator)
		: [reference.node];
	return {
		locator,
		parent: reference.parent,
		nodes,
		story: storyOf(locator),
	};
}

/** The text-box story a locator lives in (`tbx2:p0` → `tbx2`), else `body`.
 *  Cells don't start a story: moving between the body and a cell is fine. */
export function storyOf(locator: string): string {
	return /^(tbx\d+):/.exec(locator)?.[1] ?? "body";
}

/** A heading plus every following sibling up to (not including) the next
 *  heading of the same or higher level — the section `outline` nests under
 *  it, and what Word's Navigation pane drags. Stops before the body's
 *  trailing `<w:sectPr>`. */
function headingSubtree(
	heading: XmlNode,
	parent: XmlNode[],
	locator: string,
): XmlNode[] {
	const level = paragraphHeadingLevel(heading);
	if (level === null) {
		throw new MoveError(
			"USAGE",
			`--subtree needs a heading, but ${locator} isn't one (no Heading style)`,
			"Run `docx outline FILE` for the heading locators, or move a range with `--at pN-pM`.",
		);
	}
	const start = parent.indexOf(heading);
	const nodes: XmlNode[] = [heading];
	for (let index = start + 1; index < parent.length; index++) {
		const sibling = parent[index] as XmlNode;
		if (sibling.tag === "w:sectPr") break;
		const siblingLevel = paragraphHeadingLevel(sibling);
		if (siblingLevel !== null && siblingLevel <= level) break;
		nodes.push(sibling);
	}
	return nodes;
}

function paragraphHeadingLevel(node: XmlNode): number | null {
	if (node.tag !== "w:p") return null;
	const style = node
		.findChild("w:pPr")
		?.findChild("w:pStyle")
		?.getAttribute("w:val");
	return headingLevel(style);
}
