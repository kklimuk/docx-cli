import type { Document } from "../ast/document";
import {
	inheritRprChild,
	isInheritableRunProperty,
	isTypographyRunProperty,
} from "../blocks";
import { w } from "../jsx";
import {
	inheritParagraphFormattingIfPlain,
	textRuns,
} from "../paragraph-inheritance";
import { partitionParagraphRuns, type XmlNode } from "../parser";
import { PARA_MARK_ONLY_RPR_CHILDREN } from "../track-changes/preserve-formatting";

/** Re-point freshly-built list paragraphs at the HOST paragraph's list.
 *  Markdown's `- item` / `1. item` always mints a FRESH numId — a brand-new list.
 *  Dropped next to (or in place of) an existing list item, that restarts
 *  numbering at the split point and, for bullets, can switch glyph and indent.
 *  An agent writing `- New bullet` beside a bullet means "another item in THIS
 *  list", so the new items adopt the host's numId (same list kind only — a bullet
 *  list next to a numbered one is a deliberate kind change), nesting their levels
 *  under the host's. The minted numId goes unreferenced, which is harmless (see
 *  the relationship invariant: orphans are safe, dangling references are not).
 *
 *  With `copyHostLook` (insert: the new item is a SIBLING of the host) the joined
 *  paragraphs also take the host's paragraph layout (style, spacing, indent,
 *  tabs, alignment — new values win) and run typography, so a bullet inserted
 *  into a Calibri list is a Calibri bullet, not the style chain's Times New Roman.
 *  Edit doesn't need it: its own replaced-paragraph inheritance already covers a
 *  bullet retyped as a bullet.
 *
 *  Returns the paragraphs it joined to the host list. */
export function continueList(
	document: Document,
	hostParagraph: XmlNode,
	newParagraphs: XmlNode[],
	options: { copyHostLook?: boolean } = {},
): XmlNode[] {
	const host = readListContext(hostParagraph);
	if (!host) return [];
	const numbering = document.numbering;
	if (!numbering) return [];
	const hostFormat = numbering.getFormat(String(host.numId), host.level);
	// `numFmt="none"` is an unnumbered list — not something a fresh ordered/bullet
	// list should silently continue into (matches `Lists.isOrdered`, which also
	// excludes "none"). Only a real bullet or ordered host is continuable.
	if (!hostFormat || hostFormat === "none") return [];
	const hostKind = listKind(hostFormat);
	const joined: XmlNode[] = [];
	for (const paragraph of newParagraphs) {
		const fresh = readListContext(paragraph);
		if (!fresh) continue;
		if (fresh.numId !== host.numId) {
			// A bullet list beside a numbered one (or vice versa) is a deliberate
			// kind change — it keeps its own list.
			const freshKind = listKind(numbering.getFormat(String(fresh.numId), 0));
			if (freshKind !== hostKind) continue;
			repointToHost(paragraph, fresh.level, host);
		}
		joined.push(paragraph);
	}
	if (options.copyHostLook && joined.length > 0) {
		const typography = hostTypography(hostParagraph);
		for (const paragraph of joined) {
			adoptHostLayout(hostParagraph, host.level, paragraph);
			if (typography) mergeRunProperties(paragraph, typography);
		}
	}
	return joined;
}

function listKind(format: string | undefined): "bullet" | "ordered" {
	return format === "bullet" ? "bullet" : "ordered";
}

/** Point a fresh list paragraph at the host's numId, nesting its level under
 *  the host's (capped at Word's deepest level, 8). */
function repointToHost(
	paragraph: XmlNode,
	freshLevel: number,
	host: { level: number; numId: number },
): void {
	const numPr = paragraph.findChild("w:pPr")?.findChild("w:numPr");
	if (!numPr) return;
	numPr.findChild("w:numId")?.setAttribute("w:val", String(host.numId));
	const shifted = Math.min(freshLevel + host.level, 8);
	const ilvlNode = numPr.findChild("w:ilvl");
	if (ilvlNode) {
		ilvlNode.setAttribute("w:val", String(shifted));
		return;
	}
	// CT_NumPr order: <w:ilvl> precedes <w:numId>.
	if (shifted > 0) numPr.children.unshift(<w.ilvl w-val={String(shifted)} />);
}

/** The (numId, level) list membership of a paragraph, or null when it isn't a
 * (valid) list item. */
export function readListContext(
	anchor: XmlNode,
): { level: number; numId: number } | null {
	if (anchor.tag !== "w:p") return null;
	const numPr = anchor.findChild("w:pPr")?.findChild("w:numPr");
	if (!numPr) return null;
	// `numId="0"` is the OOXML sentinel for "remove this paragraph from any
	// numbered list" (ECMA-376 §17.9.18) — it's NOT a valid list to inherit.
	// `Number("")` is `0`, so guard explicitly against missing val too.
	const numIdRaw = numPr.findChild("w:numId")?.getAttribute("w:val");
	if (!numIdRaw) return null;
	const numId = Number(numIdRaw);
	if (!Number.isFinite(numId) || numId <= 0) return null;
	// A garbage `w:ilvl` (malformed doc) must not propagate NaN into a computed
	// level that a caller writes back as `w:ilvl="NaN"` — fall back to level 0.
	const levelRaw = Number(
		numPr.findChild("w:ilvl")?.getAttribute("w:val") ?? "0",
	);
	const level = Number.isFinite(levelRaw) && levelRaw >= 0 ? levelRaw : 0;
	return { level, numId };
}

/** The host item's first text run's typography (face, size, color, language,
 *  script settings — `isTypographyRunProperty`, the same set an inserted
 *  paragraph takes from its neighbor) — the list's look, minus its emphasis (a
 *  bolded lead-in word is that item's decoration, not the list's). */
function hostTypography(hostParagraph: XmlNode): XmlNode | null {
	const source = textRuns(hostParagraph)[0]?.findChild("w:rPr");
	if (!source) return null;
	const properties = source.clone();
	properties.children = properties.children.filter(
		(child) =>
			isInheritableRunProperty(child) && isTypographyRunProperty(child),
	);
	return properties.children.length > 0 ? properties : null;
}

/** Lay a joined item out like its host — what Enter at the end of the host
 *  does. The markdown walker stamps `ListParagraph` on every item; it's dropped
 *  first so the host's OWN list style (Word's `ListBullet`, or none for a Google
 *  Docs bullet) and its direct spacing/tabs come across, instead of
 *  `bringsNewBlockStructure` reading the style mismatch as new structure and
 *  copying nothing. A NESTED item keeps its numbering level's indent: the
 *  host's direct `<w:ind>` positions the host's level, and copied onto a deeper
 *  item it would un-nest it. The host's paragraph-mark revision markers
 *  (`<w:ins>`/`<w:del>` in its mark rPr) are the HOST's tracked history —
 *  cloned, they'd duplicate a revision id, mark an untracked insert as tracked,
 *  and (under tracking) give the mark two `<w:ins>`, which is schema-invalid. */
function adoptHostLayout(
	hostParagraph: XmlNode,
	hostLevel: number,
	paragraph: XmlNode,
): void {
	const own = paragraph.findChild("w:pPr");
	if (own) {
		own.children = own.children.filter((child) => child.tag !== "w:pStyle");
	}
	inheritParagraphFormattingIfPlain(hostParagraph, [paragraph], undefined);
	const pPr = paragraph.findChild("w:pPr");
	if (!pPr) return;
	if (readListContext(paragraph)?.level !== hostLevel) {
		pPr.children = pPr.children.filter((child) => child.tag !== "w:ind");
	}
	const markRpr = pPr.findChild("w:rPr");
	if (!markRpr) return;
	markRpr.children = markRpr.children.filter(
		(child) => !PARA_MARK_ONLY_RPR_CHILDREN.has(child.tag),
	);
	if (markRpr.children.length === 0) {
		pPr.children = pPr.children.filter((child) => child !== markRpr);
	}
}

/** Give each run the typography children it doesn't set itself (markdown
 *  `**bold**` keeps its `<w:b/>` and gains the list's face). A run with a
 *  CHARACTER style states part of its look through it: a `Hyperlink` run keeps
 *  its link color and an inline `Code` run its monospace face. */
function mergeRunProperties(paragraph: XmlNode, typography: XmlNode): void {
	for (const run of partitionParagraphRuns(paragraph).runs) {
		const own = run.findChild("w:rPr");
		if (!own) {
			run.children.unshift(typography.clone());
			continue;
		}
		const characterStyle = own.findChild("w:rStyle")?.getAttribute("w:val");
		for (const child of typography.children) {
			if (characterStyle && child.tag === "w:color") continue;
			if (characterStyle === "Code" && child.tag === "w:rFonts") continue;
			inheritRprChild(own, child);
		}
	}
}
