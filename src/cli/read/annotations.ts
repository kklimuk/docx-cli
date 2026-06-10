/** Shared HTML-comment annotation convention for `read --markdown`.
 *
 * `read` surfaces structural document facts the GFM body can't show as HTML
 * comments shaped `<!-- docx:TYPE [bareId] key="value" … -->` — section breaks
 * (`docx:section`, carrying an `applies-to="pX..pY (above)"` scope on deviating
 * sections so the columns/type are unambiguously tied to the content they
 * govern), page geometry (`docx:page`), table widths/borders (`docx:table`),
 * per-cell merge/shading (`docx:cell`), image size/placement (`docx:image`),
 * the document-level track-changes state (`docx:track-changes on`, at the head,
 * deviation-only), and a layout hazard (`docx:layout`) on a tab-aligned paragraph
 * inside a multi-column section, where tab stops wrap mid-line in the narrow
 * column — a render-only break Markdown can't show.
 *
 * NAMING RULE: a BARE comment is a locator (`<!-- p0 -->`, `<!-- t0:r0c0:p0 -->`)
 * — pure addressing. Anything docx-cli adds BEYOND an address is a `docx:TYPE`
 * annotation, carrying the relevant locator as a bare leading token when it has
 * one. Metadata never rides a bare locator, so `docx:` greps as "everything
 * docx-cli added" and locators stay short.
 *
 * These are read-time VISIBILITY hints, not round-trip carriers: the markdown
 * importer DROPS every one of them (`walker.tsx` returns `[]` for block `html`
 * nodes; the inline walker drops inline comments). The structure survives normal
 * edits via in-place XML mutation, `read --ast` is the lossless view, and the
 * authoring verbs (`docx sections`, `docx tables …`) manage
 * it — so a from-scratch `create` is deliberately lossy for non-Markdown
 * structure, but the agent SAW it. Re-emitted fresh on every read, so they never
 * accrete. No comment is parse-back — not even `docx:base` (the run-formatting
 * baseline note), which the importer also drops.
 *
 * Every value is escaped via `htmlAttr` so a pathological value can't close the
 * comment early (`-->`) or inject a sibling. See src/core/markdown/CLAUDE.md.
 */

export type NotePair = readonly [key: string, value: string | number];

/** Build a `<!-- docx:TYPE key="v" … -->` annotation. Callers omit pairs whose
 * value matches the document default (deviation-only — a note that just repeats
 * the default is noise, not signal). `bareTokens` ride after the type as plain
 * words (a section's `sN` locator id, ignored by the attr parser on import). */
export function formatNote(
	type: string,
	pairs: readonly NotePair[],
	bareTokens: readonly string[] = [],
): string {
	const attrs = pairs.map(([key, value]) => htmlAttr(key, String(value)));
	const parts = [`docx:${type}`, ...bareTokens, ...attrs];
	return `<!-- ${parts.join(" ")} -->`;
}

/** An HTML attribute `key="value"` with the value escaped so it can't terminate
 * the attribute/comment early or inject (`"`, `&`, `<`, `>`). Used for the
 * (dropped) `docx:` comment annotations AND for the run-formatting HTML spans
 * (`<span style>`, `<u data-underline>`, …) — for the latter, remark's HTML
 * parser decodes the entities back so the value round-trips verbatim. `&` goes
 * first so introduced entities aren't double-escaped. */
export function htmlAttr(key: string, value: string): string {
	const escaped = value
		.replace(/&/g, "&amp;")
		.replace(/</g, "&lt;")
		.replace(/>/g, "&gt;")
		.replace(/"/g, "&quot;");
	return `${key}="${escaped}"`;
}

/** Format a twip count (1/20 pt) as inches for human/agent legibility, trimming
 * trailing zeros (`1440 → "1"`, `1080 → "0.75"`). Three decimals, not two:
 * common Word margins land on 1/8-inch steps (0.625, 0.875, 0.125) that round
 * wrong at two places (0.625 → "0.63"). `read --ast` carries the exact twips. */
export function twipsToInches(twips: number): string {
	const inches = twips / 1440;
	return `${Number.parseFloat(inches.toFixed(3))}`;
}

/** Format an EMU count (914400 per inch) as inches, trimming trailing zeros.
 * Image extents (`<wp:extent>`) are in EMU; `read --ast` carries the raw EMU. */
export function emuToInches(emu: number): string {
	return `${Number.parseFloat((emu / 914400).toFixed(2))}`;
}
