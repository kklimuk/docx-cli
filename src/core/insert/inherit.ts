import {
	applyParagraphOptionsInPlace,
	ensureParagraphProperties,
	insertPprChildInOrder,
	isInheritableRunProperty,
	isTypographyRunProperty,
} from "../blocks";
import {
	ATTRIBUTE_BAG_TAGS,
	isHeadingLikeStyle,
	overlayAttributes,
	paragraphStyleId,
} from "../paragraph-inheritance";
import type { XmlNode } from "../parser";
import { paragraphMarkRunRpr } from "../track-changes/preserve-formatting";

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
	// A heading/title anchor lends nothing: inserting after it produces body
	// text, not another heading.
	if (isHeadingLikeStyle(anchorStyle)) return;
	const anchorRunProperties = firstRunProperties(anchor);
	const anchorIsList = hasListMembership(anchor);
	for (const block of blocks) {
		if (block.tag !== "w:p") continue;
		if (paragraphStyleId(block)) continue; // explicitly styled — leave it
		if (hasListMembership(block)) continue; // explicit list — leave it
		// Don't copy a list anchor's pStyle (it would carry ListParagraph without
		// the numbering) or its layout (its indent IS the list geometry); run
		// formatting still blends below.
		if (!anchorIsList) {
			if (anchorStyle) {
				applyParagraphOptionsInPlace(block.children, { style: anchorStyle });
			}
			inheritParagraphLayout(block, anchor);
		}
		if (anchorRunProperties) {
			applyRunPropertiesToBareRuns(block, anchorRunProperties);
		}
	}
}

/** The `<w:pPr>` children that make a line LAY OUT like its neighbor — what
 * Word carries into the new paragraph when you press Enter at the end of one.
 * Copied only when the new paragraph doesn't set the same child itself (an
 * explicit `--alignment`/`--indent-left`/`--tabs` still wins). Without this,
 * re-inserting a résumé entry line next to its siblings came out with no tab
 * stops and no indent, so "Harvard University\tCambridge, MA" rendered with the
 * location jammed mid-line while every other line right-aligned it. */
const LAYOUT_PPR_TAGS = ["w:jc", "w:spacing", "w:ind", "w:tabs"] as const;

function inheritParagraphLayout(block: XmlNode, anchor: XmlNode): void {
	const anchorPpr = anchor.findChild("w:pPr");
	if (!anchorPpr) return;
	const sources = LAYOUT_PPR_TAGS.map((tag) => anchorPpr.findChild(tag)).filter(
		(node): node is XmlNode => node !== undefined,
	);
	if (sources.length === 0) return;
	const pPr = ensureParagraphProperties(block);
	for (const source of sources) {
		const own = pPr.findChild(source.tag);
		if (!own) {
			insertPprChildInOrder(pPr, source.clone());
			continue;
		}
		// `<w:spacing>`/`<w:ind>` are attribute bags: an explicit `--space-before`
		// keeps the anchor's after/line, exactly as the markdown ride-along (and
		// `edit`'s pPr merge) does. Other children are one value — own wins.
		if (ATTRIBUTE_BAG_TAGS.has(source.tag)) {
			pPr.children[pPr.children.indexOf(own)] = overlayAttributes(source, own);
		}
	}
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
 * run-bearing wrappers like `<w:hyperlink>`), reduced to typography, or null
 * when the first run has nothing inheritable. An anchor with NO run lends its
 * paragraph-mark rPr WHOLE (emphasis included): that mark is what Word types
 * with in an empty paragraph, and the empty anchor is usually the very cell
 * paragraph a bare-cell insert fills — `insert --at CELL` into a bold-marked
 * blank cell must come out bold, exactly as `edit --at CELL` does. */
function firstRunProperties(paragraph: XmlNode): XmlNode | null {
	const run = firstRun(paragraph);
	const source = run?.findChild("w:rPr") ?? paragraphMarkRunRpr(paragraph);
	if (!source) return null;
	const properties = source.clone();
	properties.children = properties.children.filter(
		(child) =>
			isInheritableRunProperty(child) &&
			(!run || isTypographyRunProperty(child)),
	);
	return properties.children.length > 0 ? properties : null;
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
