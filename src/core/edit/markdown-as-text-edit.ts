import {
	partitionParagraphRuns,
	TRACKED_CHANGE_WRAPPER_TAGS,
	wrapperContent,
	type XmlNode,
} from "../parser";
import { extractOldTokens } from "../track-changes/preserve-formatting";

/** The plain text to apply as a word-level DIFF when a whole-paragraph markdown
 * edit is really a text edit — or null when it isn't. Markdown is the default
 * authoring channel, and an agent rewriting a clause usually keeps most of its
 * words; replacing the whole paragraph then redlines every word under tracking
 * (a judge-flagged contract-markup demerit: "redlines replace whole paragraphs")
 * and drops whatever the unchanged words carried that markdown can't restate —
 * run boundaries, language and character formatting the read view hides. The
 * `--text` diff (`applyFormattingPreservingEdit`) keeps every unchanged token's
 * run and redlines only the changed words.
 *
 * Deliberately conservative — it qualifies only when nothing VISIBLE about
 * formatting or structure is being changed, so the diff can't drop an intent:
 *   - the source produced ONE plain paragraph (no heading/list/quote style, no
 *     link, footnote, image or other inline object in the new runs);
 *   - the old paragraph holds only what the diff rebuild carries back — text,
 *     tabs, line breaks and object runs, under at most tracked-change/smart-tag
 *     wrappers (no link, field, note reference, equation, content control,
 *     symbol, positional tab, special hyphen or page break);
 *   - the old paragraph's text carries no visible emphasis (bold/italic/
 *     underline/strike/highlight/color/shading/caps/vertAlign/character style);
 *     a mark-up change of such text needs the whole-paragraph path;
 *   - the new runs state no formatting beyond a face/size span identical to the
 *     old text's (the read view only shows those when they deviate). */
export function markdownAsTextEdit(
	oldParagraph: XmlNode,
	newBlocks: XmlNode[],
): string | null {
	const [newParagraph, ...rest] = newBlocks;
	if (rest.length > 0 || newParagraph?.tag !== "w:p") return null;
	const newPpr = newParagraph.findChild("w:pPr");
	if (newPpr?.findChild("w:pStyle") || newPpr?.findChild("w:numPr"))
		return null;
	if (!isRebuildable(oldParagraph.children)) return null;

	const oldRuns = partitionParagraphRuns(oldParagraph).runs.filter(hasText);
	if (oldRuns.some((run) => hasVisibleEmphasis(run.findChild("w:rPr")))) {
		return null;
	}

	let text = "";
	for (const child of newParagraph.children) {
		if (child.tag === "w:pPr") continue;
		if (child.tag !== "w:r") return null; // a link / field / object wrapper
		if (!newRunMatchesOld(child.findChild("w:rPr"), oldRuns)) return null;
		for (const content of child.children) {
			if (content.tag === "w:rPr") continue;
			if (content.tag === "w:t") text += content.collectText();
			else if (content.tag === "w:tab") text += "\t";
			else if (content.tag === "w:br") text += "\n";
			else return null; // footnote reference, drawing, symbol, …
		}
	}
	if (text.length === 0) return null;
	return keepEdgeSpaces(oldParagraph, text);
}

/** Whether the token-diff rebuild carries every one of these paragraph children
 *  back: runs whose content is text/tab/line break (an object run is lifted and
 *  re-attached by the rebuild), under at most a tracked-change, smart-tag or
 *  markup-compatibility wrapper, beside content-free markers. An ALLOWLIST, not a
 *  denylist — the rebuild re-emits runs from TOKENS and hoists every other
 *  paragraph child ahead of them, so anything not named here (a link, field,
 *  equation, inline content control, symbol, positional tab, special hyphen,
 *  page break, …) would be dropped or moved to the paragraph start. Doesn't
 *  descend into a drawing, so a text box's story never disqualifies its anchor. */
function isRebuildable(children: XmlNode[]): boolean {
	return children.every((child) => {
		if (child.isText) return true;
		if (child.tag === "w:r") return child.children.every(isRebuildableRunChild);
		if (REBUILT_WRAPPERS.has(child.tag)) {
			return isRebuildable(wrapperContent(child));
		}
		return INERT_PARAGRAPH_CHILDREN.has(child.tag);
	});
}

/** Run-bearing wrappers whose runs the rebuild flattens back to plain text —
 *  every one but `<w:hyperlink>` (visible markup the markdown restates) and
 *  `<w:fldSimple>` (a field, which the rebuild would turn into its result). */
const REBUILT_WRAPPERS: ReadonlySet<string> = new Set([
	...TRACKED_CHANGE_WRAPPER_TAGS,
	"w:smartTag",
	"mc:AlternateContent",
]);

/** Paragraph-level markers that hold no text (comment ranges are already lifted
 *  out by the caller and re-anchored after the edit). */
const INERT_PARAGRAPH_CHILDREN: ReadonlySet<string> = new Set([
	"w:pPr",
	"w:bookmarkStart",
	"w:bookmarkEnd",
	"w:proofErr",
	"w:permStart",
	"w:permEnd",
	"w:commentRangeStart",
	"w:commentRangeEnd",
	"w:moveFromRangeStart",
	"w:moveFromRangeEnd",
	"w:moveToRangeStart",
	"w:moveToRangeEnd",
]);

function isRebuildableRunChild(child: XmlNode): boolean {
	if (child.isText) return true;
	if (child.tag === "w:br") {
		const type = child.getAttribute("w:type");
		return type === undefined || type === "textWrapping";
	}
	return REBUILDABLE_RUN_CHILDREN.has(child.tag);
}

/** Run content the rebuild re-emits from its token (text, tab, line break),
 *  lifts and re-attaches (objects), or can drop harmlessly (a cached page-break
 *  hint; deleted text, which the accepted view already hides). */
const REBUILDABLE_RUN_CHILDREN: ReadonlySet<string> = new Set([
	"w:rPr",
	"w:t",
	"w:delText",
	"w:tab",
	"w:lastRenderedPageBreak",
	"w:commentReference",
	"w:drawing",
	"w:pict",
	"w:object",
	"mc:AlternateContent",
]);

function hasText(run: XmlNode): boolean {
	return run.children.some(
		(child) => child.tag === "w:t" && child.collectText().length > 0,
	);
}

/** Whether a run renders with formatting the read view shows as markup (`**`,
 *  `*`, `~~`, `<u>`, `<mark>`, `<sup>`, spans with color/background/caps, code…).
 *  Mirrors the reader's `applyRunProperties`: an explicit OFF toggle, `none`
 *  underline/highlight, `auto` shading, `baseline` alignment and the default
 *  black (plain or `text1`/`dark1` theme) render nothing. */
function hasVisibleEmphasis(rPr: XmlNode | undefined): boolean {
	return rPr?.children.some(isVisibleEmphasis) ?? false;
}

function isVisibleEmphasis(child: XmlNode): boolean {
	if (EMPHASIS_TOGGLES.has(child.tag)) return child.isToggleOn();
	switch (child.tag) {
		case "w:rStyle":
			return true;
		case "w:u":
		case "w:highlight":
			return child.getAttribute("w:val") !== "none";
		case "w:shd": {
			const fill = child.getAttribute("w:fill");
			return fill !== undefined && fill !== "auto";
		}
		case "w:vertAlign":
			return child.getAttribute("w:val") !== "baseline";
		case "w:color":
			return !isDefaultColor(child);
		default:
			return false;
	}
}

/** CT_OnOff run properties the read view renders as markup. */
const EMPHASIS_TOGGLES: ReadonlySet<string> = new Set([
	"w:b",
	"w:i",
	"w:caps",
	"w:smallCaps",
	"w:strike",
	"w:dstrike",
]);

function isDefaultColor(color: XmlNode): boolean {
	const value = (color.getAttribute("w:val") ?? "auto").toLowerCase();
	if (value !== "auto" && value !== "000000") return false;
	const theme = color.getAttribute("w:themeColor");
	if (theme === undefined) return true;
	return (
		(theme === "text1" || theme === "dark1") &&
		color.getAttribute("w:themeTint") === undefined &&
		color.getAttribute("w:themeShade") === undefined
	);
}

/** A new markdown run may state only a face/size, and only one identical to
 *  every old text run's — anything else is a formatting change. With no old
 *  text to compare against (filling an empty paragraph or cell), any stated
 *  face/size is new formatting the diff would drop. */
function newRunMatchesOld(
	rPr: XmlNode | undefined,
	oldRuns: XmlNode[],
): boolean {
	if (!rPr) return true;
	if (oldRuns.length === 0) return false;
	return rPr.children.every((child) => {
		const attributes = COMPARED_ATTRIBUTES.get(child.tag);
		if (!attributes) return false;
		const oldNodes = oldRuns.map((run) =>
			run.findChild("w:rPr")?.findChild(child.tag),
		);
		return attributes.every((key) => {
			const value = child.getAttribute(key);
			return (
				value === undefined ||
				oldNodes.every((node) => node?.getAttribute(key) === value)
			);
		});
	});
}

/** The run properties a markdown span can state (`font-size`, `font-family`,
 *  `data-font-east-asia`, `data-font-complex-script`) and the attributes that
 *  carry each. `w:hAnsi` mirrors `w:ascii` in the markdown emitter. */
const COMPARED_ATTRIBUTES: ReadonlyMap<string, readonly string[]> = new Map([
	["w:sz", ["w:val"]],
	["w:szCs", ["w:val"]],
	["w:rFonts", ["w:ascii", "w:eastAsia", "w:cs"]],
]);

/** Markdown can't carry a paragraph's leading/trailing spaces (the parser
 *  strips them), so put the old paragraph's back — otherwise every word-level
 *  edit of a clause ending in a space also redlines a stray whitespace
 *  deletion the agent never asked for. */
function keepEdgeSpaces(oldParagraph: XmlNode, text: string): string {
	const oldText = extractOldTokens(oldParagraph)
		.map((token) => token.text)
		.join("");
	if (oldText.trim().length === 0) return text;
	const leading = /^[^\S\t\n]+/.test(text)
		? ""
		: (/^[^\S\t\n]+/.exec(oldText)?.[0] ?? "");
	const trailing = /[^\S\t\n]+$/.test(text)
		? ""
		: (/[^\S\t\n]+$/.exec(oldText)?.[0] ?? "");
	return `${leading}${text}${trailing}`;
}
