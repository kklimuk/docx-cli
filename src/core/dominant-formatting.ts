import type { Document } from "./ast/document";
import { flattenParagraphs } from "./ast/text";
import type { Block, TextRun } from "./ast/types";
import { inheritRprChild } from "./blocks";
import { isCodeBlockStyleId } from "./code-block/style";
import {
	isHeadingLikeStyle,
	paragraphStyleId,
	paragraphsIn,
	textRuns,
} from "./paragraph-inheritance";
import { XmlNode } from "./parser";
import { applyStylePeerFormatting, stylePeers } from "./style-peer-formatting";
import { paragraphMarkRunRpr } from "./track-changes/preserve-formatting";

/** Make freshly-authored markdown/text look like the document it lands in,
 * for whatever nothing closer (the replaced paragraph's rPr, the insert
 * anchor, a joined list) already set: first the direct formatting the new
 * paragraph's style peers agree on (`applyStylePeerFormatting`), then the
 * face/size below. The one entry point `edit` and `insert` call, so the order
 * of the fill-only-what's-unset passes lives in one place. */
export function matchDocumentLook(
	document: Document,
	newBlocks: XmlNode[],
	options: { position?: XmlNode } = {},
): void {
	applyStylePeerFormatting(document, newBlocks);
	applyDominantFormatting(document, newBlocks, options);
}

/** Give freshly-authored paragraphs the face/size of WHERE THEY LAND wherever
 * nothing closer already set it. For an edit that's the replaced paragraph's own
 * look (`position`: the face it rendered in — explicit, or through its style —
 * and its explicit size); only when that can't be determined does the new
 * content take the document's DOMINANT formatting (the face/size `read` declares
 * in its `<!-- docx:base … -->` note). A run with its own `<w:rFonts>`/`<w:sz>`
 * (the replaced paragraph's inherited rPr, the neighbor/list on insert, explicit
 * markdown spans) always keeps it. Without it, a new paragraph
 * falls back to its STYLE chain, which in real documents often isn't what's on
 * the page: the résumé's Heading1 style resolves to Times New Roman while every
 * Heading1 on the page stamps Calibri directly, so `edit --markdown "# Experience"`
 * rendered serif next to its sans-serif siblings (haiku r3).
 *
 * The reference population is the paragraph's STYLE PEERS when the document has
 * any (a new Heading1 looks like the existing Heading1s — so a Cambria-heading /
 * Calibri-body theme keeps Cambria headings), else the whole document (the
 * `docx:base` baseline). Size is only taken from the whole-document baseline for
 * body text — a heading with no peers keeps its style's size rather than being
 * shrunk to body size. Code blocks, inline code, and
 * `*Reference` runs keep their style's font. */
function applyDominantFormatting(
	document: Document,
	newBlocks: XmlNode[],
	{ position }: { position?: XmlNode } = {},
): void {
	const styles = document.styles;
	const resolve: FontResolver | undefined = styles?.resolveFont.bind(styles);
	const existing = document.body.blocks;
	const inheritedFonts = resolveInheritedFonts(existing, resolve);
	const documentBaseline = detectFormatBaseline(existing, inheritedFonts);
	// The face the document's runs literally STATE (no style resolution) — what
	// makes a theme-resolved face safe to override (below).
	const statedFont = detectFormatBaseline(existing, undefined).font;
	const positionLook = position ? lookAt(position, resolve) : null;
	const peerBaselines = new Map<string, RunFormatBaseline | null>();
	const baselineFor = (styleId: string): RunFormatBaseline | null => {
		if (!peerBaselines.has(styleId)) {
			const peers = stylePeers(document, styleId);
			peerBaselines.set(
				styleId,
				peers.length > 0 ? detectFormatBaseline(peers, inheritedFonts) : null,
			);
		}
		return peerBaselines.get(styleId) ?? null;
	};

	for (const paragraph of paragraphsIn(newBlocks)) {
		const styleId = paragraphStyleId(paragraph);
		// A code block's monospace IS its style — never body-font it.
		if (isCodeBlockStyleId(styleId)) continue;
		const peers = styleId ? baselineFor(styleId) : null;
		const fallback = peers ?? documentBaseline;
		// A theme-resolved face is only overridden for a face someone STATED —
		// the replaced run's own rFonts, style peers, or the doc's explicit
		// majority. Each source vouches only for ITS face: peers vouch for the
		// peers' face, never for a position face that merely resolves through
		// docDefaults (every plain template body line), so that one doesn't
		// override a theme heading font.
		const face =
			positionLook?.font !== undefined
				? {
						font: positionLook.font,
						vouched:
							positionLook.fontExplicit || statedFont === positionLook.font,
					}
				: {
						font: fallback.font,
						vouched: peers !== null || statedFont === fallback.font,
					};
		const sizeHalfPoints = targetSize(styleId, positionLook, peers, fallback);
		for (const run of textRuns(paragraph)) {
			const runStyle = run
				.findChild("w:rPr")
				?.findChild("w:rStyle")
				?.getAttribute("w:val");
			if (runStyle && OWN_FONT_RUN_STYLE.test(runStyle)) continue;
			stampRun(run, {
				font:
					face.font &&
					needsFont(run, resolve?.(styleId, runStyle), face.font, face.vouched)
						? face.font
						: undefined,
				sizeHalfPoints,
			});
		}
	}
}

/** The size a new paragraph takes. A heading keeps its style's size unless its
 * own peers or the replaced heading say otherwise — an 8pt form line turned
 * `## Heading` must not render at 8pt, but a new heading next to Heading1s that
 * all stamp 14pt is 14pt. */
function targetSize(
	styleId: string | undefined,
	positionLook: PositionLook | null,
	peers: RunFormatBaseline | null,
	fallback: RunFormatBaseline,
): number | undefined {
	const headingLike = isHeadingLikeStyle(styleId);
	if (positionLook && (!headingLike || positionLook.styleId === styleId)) {
		return positionLook.sizeHalfPoints;
	}
	if (peers !== null || !headingLike) return fallback.sizeHalfPoints;
	return undefined;
}

/** The look a paragraph HAD: the face its first text run rendered in (explicit
 * `<w:rFonts>`, else resolved through its paragraph/character style — an empty
 * paragraph answers from its paragraph mark) and its explicit size. A size that
 * came from the style isn't stamped: the same style on the new paragraph already
 * reproduces it. */
function lookAt(
	paragraph: XmlNode,
	resolve: FontResolver | undefined,
): PositionLook {
	const styleId = paragraphStyleId(paragraph);
	const firstText = textRuns(paragraph)[0];
	const rPr = firstText
		? firstText.findChild("w:rPr")
		: paragraphMarkRunRpr(paragraph);
	const runStyle = rPr?.findChild("w:rStyle")?.getAttribute("w:val");
	const size = Number(rPr?.findChild("w:sz")?.getAttribute("w:val"));
	const stated = rPr?.findChild("w:rFonts")?.getAttribute("w:ascii");
	return {
		styleId,
		font: stated ?? resolve?.(styleId, runStyle),
		fontExplicit: stated !== undefined,
		sizeHalfPoints: Number.isFinite(size) && size > 0 ? size : undefined,
	};
}

/** Character styles that ARE a font statement (`Code` → monospace; footnote /
 * comment `*Reference` marks) — never restamp their face. */
const OWN_FONT_RUN_STYLE = /^code$|reference$/i;

/** Whether a run without an explicit face renders in something other than the
 * target face. An unresolvable (theme) face is only overridden when the target
 * is VOUCHED FOR — by style peers, or by the document's runs stating that face
 * outright (an all-Arial form). A bare template doc's Calibri target is itself
 * just the docDefaults resolution, so a theme heading font (Calibri Light) there
 * is the document's real design and stays. */
function needsFont(
	run: XmlNode,
	resolvedFace: string | undefined,
	targetFace: string,
	vouched: boolean,
): boolean {
	// A Latin face the run states itself wins; an rFonts carrying only other
	// script slots (the read view's `data-font-complex-script` span) doesn't.
	const fonts = run.findChild("w:rPr")?.findChild("w:rFonts");
	if (
		fonts?.getAttribute("w:ascii") !== undefined ||
		fonts?.getAttribute("w:asciiTheme") !== undefined
	) {
		return false;
	}
	if (resolvedFace === undefined) return vouched;
	return resolvedFace !== targetFace;
}

/** Add the missing `<w:rFonts>`/`<w:sz>` to one run, never overriding what the
 * run already states (an rFonts carrying only other script slots gains the
 * Latin ones). The face covers the Latin slots (ASCII, high-ANSI) — the
 * ones `docx:base` describes; stamping complex-script too would surface as a
 * spurious `data-font-complex-script` span on read. */
function stampRun(
	run: XmlNode,
	{ font, sizeHalfPoints }: { font?: string; sizeHalfPoints?: number },
): void {
	const additions: XmlNode[] = [];
	const own = run.findChild("w:rPr");
	if (font) {
		additions.push(
			new XmlNode("w:rFonts", {
				"w:ascii": font,
				"w:hAnsi": font,
			}),
		);
	}
	// Size travels as a pair: a run that states `<w:sz>` keeps both halves.
	if (sizeHalfPoints !== undefined && !own?.findChild("w:sz")) {
		additions.push(new XmlNode("w:sz", { "w:val": String(sizeHalfPoints) }));
		additions.push(new XmlNode("w:szCs", { "w:val": String(sizeHalfPoints) }));
	}
	if (additions.length === 0) return;
	const rPr = own ?? new XmlNode("w:rPr");
	if (!own) run.children.unshift(rPr);
	// Merges into an existing rFonts slot-by-slot (a cs-only rFonts gains its
	// Latin face) and never overrides a child the run already has.
	for (const child of additions) inheritRprChild(rPr, child);
}

/** Resolve the inherited face of every bare text run once, so the baseline
 *  majority and the per-run spans agree on what each run renders in. Memoized
 *  per (paragraph style, character style) pair — a document has a handful of
 *  those and thousands of runs, and each resolve walks the styles part. */
export function resolveInheritedFonts(
	blocks: Block[],
	resolve: FontResolver | undefined,
): WeakMap<TextRun, string> | undefined {
	if (!resolve) return undefined;
	const inherited = new WeakMap<TextRun, string>();
	const faceByStyles = new Map<string, string | undefined>();
	for (const paragraph of flattenParagraphs(blocks)) {
		for (const run of paragraph.runs) {
			if (run.type !== "text" || run.font) continue;
			// A character style Markdown renders NATIVELY (`Code` → backticks) is
			// its own font statement; resolving its Consolas into a `font-family`
			// span would double-mark every inline code run. Resolve through the
			// paragraph for those.
			const runStyle = run.runStyle === "Code" ? undefined : run.runStyle;
			const key = `${paragraph.style ?? ""}\u0000${runStyle ?? ""}`;
			if (!faceByStyles.has(key)) {
				faceByStyles.set(key, resolve(paragraph.style, runStyle));
			}
			const face = faceByStyles.get(key);
			if (face) inherited.set(run, face);
		}
	}
	return inherited;
}

/** The dominant font and size across the blocks' text (body + table runs), each
 *  reported only when it covers a MAJORITY of the characters. A clear majority
 *  is what makes omitting it from every run legible rather than lossy; below the
 *  threshold there is no single baseline and every run keeps its formatting. The
 *  font majority counts what each run RENDERS in (explicit face, else the
 *  inherited one from `inheritedFonts`) — the dominant font on the page, not the
 *  most-stamped `<w:rFonts>`; pass `undefined` to count explicit faces only. */
export function detectFormatBaseline(
	blocks: Block[],
	inheritedFonts: WeakMap<TextRun, string> | undefined,
): RunFormatBaseline {
	const fontChars = new Map<string, number>();
	const sizeChars = new Map<number, number>();
	let total = 0;
	for (const paragraph of flattenParagraphs(blocks)) {
		for (const run of paragraph.runs) {
			if (run.type !== "text") continue;
			const length = run.text.length;
			if (length === 0) continue;
			total += length;
			const font = effectiveFont(run, { inheritedFonts });
			if (font) fontChars.set(font, (fontChars.get(font) ?? 0) + length);
			if (run.sizeHalfPoints !== undefined) {
				sizeChars.set(
					run.sizeHalfPoints,
					(sizeChars.get(run.sizeHalfPoints) ?? 0) + length,
				);
			}
		}
	}
	if (total === 0) return {};
	return {
		font: majorityKey(fontChars, total),
		sizeHalfPoints: majorityKey(sizeChars, total),
	};
}

/** The face a run renders in: its explicit `font`, else the one it inherits
 *  from its paragraph/character style (undefined when unknown). */
export function effectiveFont(
	run: TextRun,
	baseline: RunFormatBaseline,
): string | undefined {
	return run.font ?? baseline.inheritedFonts?.get(run);
}

/** The map key whose accumulated weight exceeds half the total, or undefined
 * when no single key does. */
function majorityKey<K>(counts: Map<K, number>, total: number): K | undefined {
	for (const [key, weight] of counts) {
		if (weight * 2 > total) return key;
	}
	return undefined;
}

/** Document-wide run formatting so ubiquitous it reads as noise: the dominant
 *  font and size across all body + table runs. `read` emits it once as a
 *  `<!-- docx:base … -->` note and omits it from every matching run, so the body
 *  reads clean AND an agent can see the doc's baseline to match new content;
 *  `applyDominantFormatting` stamps it on new content so what the agent sees is
 *  what it gets. A value only becomes a baseline when it covers a majority of
 *  the text — otherwise every run keeps its explicit formatting. */
export type RunFormatBaseline = {
	font?: string;
	sizeHalfPoints?: number;
	/** Per-run INHERITED face for runs with no explicit `font` (resolved through
	 *  the style chain once, up front). Rides the baseline because every font
	 *  comparison already receives it; `effectiveFont` is the one accessor.
	 *  Absent when the caller can't resolve (no styles part). */
	inheritedFonts?: WeakMap<TextRun, string>;
};

/** The face a run WITHOUT explicit `<w:rFonts>` renders in, from the style
 *  chain + docDefaults (`StylesView.resolveFont`). Undefined = unknown. */
export type FontResolver = (
	paragraphStyle: string | undefined,
	runStyle: string | undefined,
) => string | undefined;

/** The look at the replaced paragraph's position (`lookAt`). */
type PositionLook = {
	font?: string;
	fontExplicit: boolean;
	sizeHalfPoints?: number;
	styleId?: string;
};
