import { applyParagraphOptionsInPlace } from "../blocks";
import type { XmlNode } from "../parser";

/** Paragraph styles whose look must NOT bleed onto inserted body content:
 * inserting after a heading/title should produce body text, not another
 * heading. Matched case-insensitively against the prefix of the pStyle id. */
const HEADING_LIKE = /^(heading[1-9]|title|subtitle|toc)/i;

/** Make freshly-inserted plain content blend into the document it lands in by
 * copying the anchor paragraph's run formatting — and paragraph style, when
 * safe — onto new runs/paragraphs that didn't bring their own. This is the
 * "inherit from neighbor" contract: `insert --after pN` of plain text in an
 * all-Arial-8pt document comes out Arial 8pt, not the bare document default
 * (verified gap). It mirrors how `edit` already inherits a paragraph's rPr when
 * it rewrites the paragraph in place.
 *
 * Deliberately conservative — it never overrides content that specified its own
 * formatting, and it bails on a heading/list anchor so it can't promote inserted
 * body text into a heading or graft a heading's size onto it:
 *   - anchor is heading/title  → inherit nothing (body-after-heading stays body)
 *   - block has its own pStyle  → leave that block entirely (e.g. a `# heading`)
 *   - block is its own list item → leave it (explicit `--list` / numPr)
 *   - a run already has `<w:rPr>` → keep it (markdown bold/links/explicit color)
 * Only applies to text-bearing inserts (the lens gates the call); structural
 * inserts (table/image/section/break/code/equation) never blend. */
export function inheritFormattingFromAnchor(
	blocks: XmlNode[],
	anchor: XmlNode,
): void {
	if (anchor.tag !== "w:p") return;
	const anchorStyle = paragraphStyleId(anchor);
	if (anchorStyle && HEADING_LIKE.test(anchorStyle)) return;
	const anchorRunProperties = firstRunProperties(anchor);
	const anchorIsList = hasListMembership(anchor);
	for (const block of blocks) {
		if (block.tag !== "w:p") continue;
		if (paragraphStyleId(block)) continue; // explicitly styled — leave it
		if (hasListMembership(block)) continue; // explicit list — leave it
		// Don't copy a list anchor's pStyle (it would carry ListParagraph without
		// the numbering); run formatting still blends below.
		if (!anchorIsList && anchorStyle) {
			applyParagraphOptionsInPlace(block.children, { style: anchorStyle });
		}
		if (anchorRunProperties) {
			applyRunPropertiesToBareRuns(block, anchorRunProperties);
		}
	}
}

function paragraphStyleId(paragraph: XmlNode): string | undefined {
	return paragraph
		.findChild("w:pPr")
		?.findChild("w:pStyle")
		?.getAttribute("w:val");
}

/** True when the paragraph belongs to a real numbered/bulleted list — a
 * `<w:numPr>` with a positive `numId` (id 0 is the OOXML "remove from list"
 * sentinel, not a list to preserve). */
function hasListMembership(paragraph: XmlNode): boolean {
	const numId = paragraph
		.findChild("w:pPr")
		?.findChild("w:numPr")
		?.findChild("w:numId")
		?.getAttribute("w:val");
	return numId !== undefined && Number(numId) > 0;
}

/** A clone of the first run's `<w:rPr>` in the paragraph (descending into
 * run-bearing wrappers like `<w:hyperlink>`), or null when the first run has no
 * explicit formatting to inherit. */
function firstRunProperties(paragraph: XmlNode): XmlNode | null {
	const run = firstRun(paragraph);
	const properties = run?.findChild("w:rPr");
	return properties ? properties.clone() : null;
}

function firstRun(node: XmlNode): XmlNode | undefined {
	for (const child of node.children) {
		if (child.tag === "w:r") {
			// Skip a comment/footnote/endnote-reference run: its rPr is the
			// CommentReference/FootnoteReference character style, which must not
			// bleed onto inserted plain text. Inherit from the first REAL text run.
			if (isReferenceRun(child)) continue;
			return child;
		}
		const nested = firstRun(child);
		if (nested) return nested;
	}
	return undefined;
}

/** A run that carries a comment/footnote/endnote reference (by element or by a
 *  `*Reference` character style) rather than authored text. */
function isReferenceRun(run: XmlNode): boolean {
	for (const child of run.children) {
		if (
			child.tag === "w:commentReference" ||
			child.tag === "w:footnoteReference" ||
			child.tag === "w:endnoteReference"
		) {
			return true;
		}
	}
	const rStyle = run
		.findChild("w:rPr")
		?.findChild("w:rStyle")
		?.getAttribute("w:val");
	return rStyle !== undefined && /Reference$/i.test(rStyle);
}

/** Give every run in the paragraph that has no `<w:rPr>` of its own a clone of
 * the inherited properties, so plain inserted runs adopt the surrounding
 * formatting. Runs that already carry properties (markdown bold, hyperlink
 * styling, an explicit color) keep theirs untouched. */
function applyRunPropertiesToBareRuns(
	node: XmlNode,
	properties: XmlNode,
): void {
	for (const child of node.children) {
		if (child.tag === "w:r") {
			if (!child.findChild("w:rPr")) child.children.unshift(properties.clone());
			continue; // never descend into a run
		}
		if (child.children.length > 0)
			applyRunPropertiesToBareRuns(child, properties);
	}
}
