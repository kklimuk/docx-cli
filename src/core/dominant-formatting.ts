import { flattenParagraphs } from "./ast/text";
import type { Block, TextRun } from "./ast/types";

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
 *  reads clean AND an agent can see the doc's baseline to match new content. A
 *  value only becomes a baseline when it covers a majority of the text —
 *  otherwise every run keeps its explicit formatting. */
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
