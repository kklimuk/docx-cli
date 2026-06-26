export type DocProperties = {
	title?: string;
	author?: string;
	created?: string;
	modified?: string;
};

export type Block = Paragraph | Table | SectionBreak;

export type Paragraph = {
	id: string;
	type: "paragraph";
	style?: string;
	alignment?: "left" | "center" | "right" | "justify";
	list?: { level: number; numId: string; ordered?: boolean };
	/** GFM task-list marker. Set when the paragraph's first content is a Word
	 * checkbox content control (`<w:sdt><w14:checkbox/></w:sdt>`); the SDT itself
	 * plus the leading whitespace run after it are stripped from `runs` so the
	 * AST carries only the task text. Markdown render emits `- [ ]` / `- [x]`. */
	taskState?: "checked" | "unchecked";
	/** Markdown blockquote depth (1 = single `>`, 2 = `> >`, etc.). Detected
	 * by the AST reader from `pStyle="Quote"` / `pStyle="QuoteListParagraph"`
	 * combined with the paragraph's `<w:ind w:left>` value (each 720-twip
	 * step is one nesting level). Markdown render prepends one `> ` per
	 * depth. Code blocks, tables, math, headings, and HRs inside a markdown
	 * blockquote intentionally do **not** carry this flag — the walker emits
	 * them at top level, breaking the quote at that point. See
	 * [src/core/markdown/CLAUDE.md](../markdown/CLAUDE.md). */
	quoteDepth?: number;
	/** Explicit tab stops from `<w:pPr><w:tabs>`, each `{ align, pos }` (pos in
	 * twips). Surfaced so `read` can flag a fragile right-alignment: a LEFT (or
	 * center) tab stop near the right margin pushes content rightward but — unlike
	 * a RIGHT tab — wraps anything wider than the gap to the margin (the résumé
	 * "San / Francisco, CA" break). Present only when the paragraph declares
	 * explicit tab stops. */
	tabStops?: { align: string; pos: number }[];
	/** Direct paragraph spacing from `<w:pPr><w:spacing>`. Present only when the
	 * paragraph sets spacing directly (not inherited from its style/the doc
	 * default). */
	spacing?: ParagraphSpacing;
	/** Direct paragraph indentation from `<w:pPr><w:ind>`. Surfaced only for
	 * non-list, non-quote paragraphs — a list/quote's `<w:ind>` is structural
	 * positioning, not direct formatting. */
	indent?: ParagraphIndent;
	runs: Run[];
};

/** `<w:pPr><w:spacing>`. `before`/`after` are twips (1/20 pt). `line` is the
 * line-spacing value: with `lineRule="auto"` it's 240ths of a line (240 =
 * single, 360 = 1.5×, 480 = double); with `exact`/`atLeast` it's twips. */
export type ParagraphSpacing = {
	before?: number;
	after?: number;
	line?: number;
	lineRule?: "auto" | "exact" | "atLeast";
};

/** `<w:pPr><w:ind>`, in twips (1440/inch). `firstLine` and `hanging` are
 * mutually exclusive (same slot, opposite sign). */
export type ParagraphIndent = {
	left?: number;
	right?: number;
	firstLine?: number;
	hanging?: number;
};

export type Table = {
	id: string;
	type: "table";
	/** Column widths in twips, extracted from <w:tblGrid><w:gridCol w:w="…"/>.
	 * Length matches the number of grid columns. May be empty if the source
	 * doc omits the tblGrid (rare but legal). */
	grid: number[];
	/** Table-level width from <w:tblPr><w:tblW/>. Optional — absent means
	 * "auto" (sum of grid). `unit` matches OOXML's w:type attribute. */
	width?: TableWidth;
	/** Summary of `<w:tblPr><w:tblBorders>` — the dominant edge style across the
	 * six edges (`"single"` / `"double"` / `"none"` / …), or `"mixed"` when edges
	 * differ. Present only when the table declares explicit borders. A SUMMARY for
	 * visibility (GFM shows none); full per-edge fidelity stays in the XML via
	 * in-place mutation, and `docx tables borders` is the way to set them. */
	borders?: string;
	/** Table style applied via `<w:tblPr><w:tblStyle w:val="…"/>` (e.g.
	 * "TableGrid") — a reference into styles.xml, distinct from the inline
	 * `borders` summary. Surfaced so `styles --used` can report the table styles
	 * a document actually applies. Present only when the table references one. */
	style?: string;
	/** Justification of the whole table on the page, from `<w:tblPr><w:jc w:val="…"/>`
	 * (`"center"` / `"right"`). Absent ⇒ the default `"left"`. Authorable via
	 * `docx tables format --at tN --align`. */
	align?: "left" | "center" | "right";
	rows: TableRow[];
};

export type TableRow = {
	cells: TableCell[];
	/** Tracked row insertion/deletion from <w:trPr><w:ins/> or <w:del/>
	 * (kind "rowIns" / "rowDel"). Present only under track-changes. */
	trackedChange?: TrackedChange;
	/** Row height from `<w:trPr><w:trHeight w:val="…" w:hRule="…"/>`. `value` is
	 * in twips; `rule` is `atLeast` (minimum, the default), `exact` (fixed), or
	 * `auto` (fit content). Present only when the row declares an explicit height.
	 * Authorable via `docx tables format --at tN:rR --row-height`. */
	height?: { value: number; rule: "atLeast" | "exact" | "auto" };
	/** Whether the row repeats as a header at the top of each page that the table
	 * spans, from `<w:trPr><w:tblHeader/>`. Present (true) only when the marker is
	 * set. Authorable via `docx tables format --at tN:rR --repeat-header`. */
	repeatHeader?: boolean;
};

export type TableCell = {
	blocks: Block[];
	/** Tracked cell insertion/deletion from <w:tcPr><w:cellIns/> or
	 * <w:cellDel/> (kind "cellIns" / "cellDel") — how column insert/delete is
	 * recorded under track-changes (one marker per cell of the column). */
	trackedChange?: TrackedChange;
	/** Horizontal merge: this cell spans N grid columns (default 1). From
	 * <w:tcPr><w:gridSpan w:val="N"/>. */
	gridSpan?: number;
	/** Vertical merge marker. From <w:tcPr><w:vMerge w:val="restart"/> or
	 * a bare <w:vMerge/> (interpreted as "continue"). Together with the cell
	 * directly above's "restart" marker, this binds adjacent rows into a
	 * single visual cell. */
	vMerge?: "restart" | "continue";
	/** Cell-level width override from <w:tcPr><w:tcW/>. Falls back to the
	 * grid column's width if absent. */
	width?: TableWidth;
	/** Cell background fill hex from `<w:tcPr><w:shd w:fill="…"/>` (6-digit,
	 * uppercase), when set to a real color (not `auto`). Surfaced as a read-time
	 * hint — GFM can't show cell shading; the fill survives edits via in-place
	 * mutation. */
	shading?: string;
	/** Vertical alignment of the cell's content from `<w:tcPr><w:vAlign w:val="…"/>`
	 * (`"center"` / `"bottom"`). Absent ⇒ the default `"top"`. Authorable via
	 * `docx tables format --at … --valign`. (Horizontal alignment is NOT a cell
	 * property — it lives on each paragraph's `<w:jc>`, surfaced as `docx:p align`.) */
	vAlign?: "top" | "center" | "bottom";
	/** Summary of `<w:tcPr><w:tcBorders>` — the set edges and their dominant style,
	 * as a compact `side:style` string (e.g. `"bottom:double"`, `"all:single"`).
	 * A read-time hint mirroring `Table.borders`; full per-edge fidelity stays in
	 * the XML via in-place mutation. Authorable via `docx tables format --at …
	 * --cell-borders`. Present only when the cell declares explicit borders. */
	borders?: string;
};

export type TableWidth = {
	value: number;
	/** Mirrors OOXML's w:type: dxa = twips (1/20 pt), pct = fiftieths of a
	 * percent (so 5000 = 100%), auto = compute from content, nil = no width. */
	unit: "dxa" | "pct" | "auto" | "nil";
};

export type SectionBreak = {
	id: string;
	type: "sectionBreak";
	columns?: number;
	sectionType?: SectionType;
	/** Page geometry from `<w:sectPr><w:pgSz>` / `<w:pgMar>`, in twips (1/20 pt
	 * — the OOXML native unit, matching the table grid). Present only when the
	 * sectPr declares the attribute. The trailing (mandatory) section break
	 * carries the document-wide geometry; an inline sectPr usually omits these
	 * (it inherits). `read --markdown` surfaces deviations from the canonical
	 * default (US Letter portrait, 1″ margins) as a leading `<!-- docx:page -->`
	 * note; `read --ast` carries the exact twips. Margins may be negative. */
	pageWidth?: number;
	pageHeight?: number;
	pageOrientation?: "portrait" | "landscape";
	marginTop?: number;
	marginRight?: number;
	marginBottom?: number;
	marginLeft?: number;
};

export type SectionType =
	| "continuous"
	| "nextPage"
	| "evenPage"
	| "oddPage"
	| "nextColumn";

export type Run =
	| TextRun
	| ImageRun
	| BreakRun
	| TabRun
	| EquationRun
	| NoteRefRun
	| ChartRun;

export type TextRun = {
	type: "text";
	text: string;
	color?: string;
	/** Theme-color reference + tint/shade modifiers from `<w:color>`, kept as raw
	 *  hex/token strings so they round-trip byte-exact alongside the `color`
	 *  fallback hex. */
	colorTheme?: string;
	colorThemeTint?: string;
	colorThemeShade?: string;
	highlight?: string;
	/** Run shading fill (`<w:shd w:fill>`) — arbitrary hex background, distinct
	 *  from the 16-name `highlight` enum. */
	shade?: string;
	bold?: boolean;
	italic?: boolean;
	underline?: string;
	/** Underline color (`<w:u w:color>`), when not the default `auto`. */
	underlineColor?: string;
	strike?: boolean;
	/** Vertical alignment (`<w:vertAlign>`): "superscript" | "subscript". */
	vertAlign?: string;
	/** Small-caps (`<w:smallCaps>`) and all-caps (`<w:caps>`) effects. */
	smallCaps?: boolean;
	allCaps?: boolean;
	font?: string;
	sizeHalfPoints?: number;
	/** Character style applied via <w:rStyle> (e.g., "Code" for inline code,
	 *  "Hyperlink", or any user-defined style). Distinct from direct formatting
	 *  (bold/italic/color/…): a character style is a reference into styles.xml
	 *  and gets provisioned via `ensureStyle` when emitted. */
	runStyle?: string;
	comments?: string[];
	trackedChange?: TrackedChange;
	hyperlink?: Hyperlink;
};

export type Hyperlink = {
	id: string;
	url?: string;
	anchor?: string;
	tooltip?: string;
};

export type ImageRun = {
	type: "image";
	id: string;
	contentType: string;
	hash: string;
	widthEmu?: number;
	heightEmu?: number;
	alt?: string;
	/** True when the drawing is `<wp:anchor>` (floating — positioned out of the
	 * text flow with wrap) rather than `<wp:inline>` (the default, flows with the
	 * text). `read --markdown` surfaces this; GFM can only show the inline shape. */
	floating?: boolean;
	/** Text-wrap mode of a floating image, from the `<wp:wrap*>` child:
	 * `square` / `tight` / `through` / `topAndBottom` / `none`. Absent for inline
	 * images. */
	wrap?: string;
	/** Horizontal placement of a floating image from `<wp:positionH>`:
	 * `left` / `center` / `right` (a `<wp:align>`), or `absolute` (a fixed
	 * `<wp:posOffset>`). Absent for inline images (they flow with the text). */
	align?: string;
	trackedChange?: TrackedChange;
};

export type BreakRun = {
	type: "break";
	kind: "page" | "line" | "column";
};

export type TabRun = {
	type: "tab";
};

/** A math equation. `latex` is reconstructed by walking the underlying
 *  `<m:oMath>` subtree (one handler per OMML element — see `@core/equation`)
 *  and is the user-facing form rendered as `$…$` / `$$…$$` in markdown.
 *  `text` is the legacy plaintext concatenation, kept as a defensive fallback
 *  for OMML constructs the walker doesn't handle cleanly (rare in practice;
 *  the walker covers the common top-80% and falls back per-subtree). `display`
 *  distinguishes inline (`<m:oMath>`) from block-level (`<m:oMathPara>`). The
 *  `id` (`eqN`) addresses the equation in document order — same scheme as
 *  imgN/linkN — so callers can target one equation via `--at eqN`. The
 *  original `<m:oMath>` XmlNode isn't on this type (not JSON-serializable);
 *  the reader stashes it in `Document.equationReferences` for emit-back paths. */
export type EquationRun = {
	type: "equation";
	id: string;
	latex: string;
	text: string;
	display: boolean;
	/** Comment IDs whose range covers this equation. Populated when
	 *  `<w:commentRangeStart>` opens before the `<m:oMath>` and closes after
	 *  — typically from the audit-comment fallback for tracked equation
	 *  edits (Word doesn't have a `<w:oMathChange>` element). */
	comments?: string[];
};

/** Reference to a footnote or endnote whose body lives in Body.footnotes or
 * Body.endnotes. The id matches the footnote/endnote id (`fn1`, `en1`). */
export type NoteRefRun = {
	type: "noteRef";
	kind: "footnote" | "endnote";
	id: string;
};

/** Placeholder for a non-picture drawing (chart, shape, SmartArt, etc.). The
 * full content is preserved in the underlying XML; the run is a marker so
 * callers can know "something visual lives here." */
export type ChartRun = {
	type: "chart";
	kind: "chart" | "shape" | "smartart" | "drawing";
};

export type Note = {
	id: string;
	text: string;
};

/** A header or footer, surfaced on `Body.headers` / `Body.footers`. One entry per
 *  `<w:sectPr>` reference, so a document-wide header (one shared part referenced
 *  from every section) appears once per section. `id` (`hdr0` / `ftr0`) is the
 *  positional handle; `sectionId` is the owning section's `sN`; `type` is the
 *  placement (`default` = all/odd pages, `first`, `even`); `text` renders fields
 *  as `{page}` / `{pages}` / `{date}` / `{styleref:NAME}` / `{filename}` / `{title}`
 *  / `{author}` tokens (the cached field value is not shown). Authored via
 *  `docx headers`/`docx footers` (the `--at sN --type T` address) — `read --markdown`
 *  surfaces it deviation-only as a `<!-- docx:header … -->` hint the importer drops. */
export type Marginal = {
	id: string;
	kind: "header" | "footer";
	type: "default" | "first" | "even";
	sectionId: string;
	text: string;
};

export type TrackedChange = {
	id: string;
	kind: TrackedChangeKind;
	author: string;
	date: string;
	revisionId: string;
};

/** OOXML revision-tracking wrappers we surface in the AST.
 *  - `ins` / `del`: <w:ins> / <w:del> — inserted / deleted runs.
 *  - `moveFrom` / `moveTo`: <w:moveFrom> / <w:moveTo> — origin / destination
 *    of a tracked move. moveFrom behaves like a delete (text leaves this
 *    location, stored as <w:delText> internally); moveTo behaves like an
 *    insert (text arrives at this location).
 *  - `sectPrChange`: <w:sectPrChange> — section-property revision marker
 *    embedded inside a <w:sectPr>. Carries a snapshot of the prior section
 *    properties (e.g. columns / type) so accept/reject can drop or restore
 *    them. Has no run text.
 *  - `pPrChange`: <w:pPrChange> — paragraph-property revision marker embedded
 *    inside a <w:pPr> (its last child). Carries a snapshot of the prior <w:pPr>
 *    (spacing / indent / style / alignment / …) so accept drops the marker
 *    (keeping the new props) and reject restores the prior pPr. The paragraph
 *    analog of `sectPrChange`; how `edit`'s paragraph-property changes are
 *    tracked. Has no run text. Empirically matches Word for Mac's shape.
 *  - `rowIns` / `rowDel`: <w:trPr><w:ins/> / <w:del/> — a whole table row
 *    inserted / deleted under tracking. Accept/reject acts on the entire
 *    <w:tr>. Has no run text of its own.
 *  - `cellIns` / `cellDel`: <w:tcPr><w:cellIns/> / <w:cellDel/> — a single
 *    cell inserted / deleted; how column insert/delete is tracked (one marker
 *    per cell of the column). Has no run text of its own.
 *  - `tblGridChange`: <w:tblGrid><w:tblGridChange> — a revision to the table
 *    grid (column widths/count), carrying a snapshot of the prior <w:tblGrid>.
 *    Emitted alongside the per-cell markers on a tracked column insertion so
 *    the width change is reversible. Mirrors `sectPrChange`. No run text.
 *  - `tblPrChange`: <w:tblPr><w:tblPrChange> — a revision to table-level
 *    properties (borders, layout…), carrying a snapshot of the prior
 *    <w:tblPr>. Mirrors `sectPrChange`. No run text.
 *  - `tcPrChange`: <w:tcPr><w:tcPrChange> — a revision to a cell's properties
 *    (gridSpan / hMerge / vMerge / width…), carrying a snapshot of the prior
 *    <w:tcPr>. How merge / unmerge is tracked: the cell-preserving merge
 *    markers stay, and reject restores the pre-merge <w:tcPr>. No run text.
 *  - `checkboxToggle`: a tracked toggle of a `<w14:checkbox>` content control's
 *    state. Word/LibreOffice emit this as an `<w:ins>` (new glyph ☒ or ☐) and
 *    `<w:del>` (old glyph ☐ or ☒) pair INSIDE `<w:sdtContent>`, plus an
 *    in-place flip of the `w14:checked` attribute (no separate
 *    `<w14:checkedChange>` element exists in the spec). We surface the pair
 *    as a single revision so `track-changes accept/reject` can keep the glyph
 *    and attribute consistent — reject infers the prior attribute value from
 *    the deleted glyph (☐ → "0", ☒ → "1") and flips the attribute back.
 */
export type TrackedChangeKind =
	| "ins"
	| "del"
	| "moveFrom"
	| "moveTo"
	| "sectPrChange"
	| "pPrChange"
	| "rowIns"
	| "rowDel"
	| "cellIns"
	| "cellDel"
	| "tblGridChange"
	| "tblPrChange"
	| "tcPrChange"
	| "checkboxToggle";

export type Comment = {
	id: string;
	author: string;
	initials?: string;
	date: string;
	text: string;
	parentId?: string;
	resolved?: boolean;
	anchor: CommentAnchor;
};

export type CommentAnchor = {
	startBlockId: string;
	startOffset: number;
	endBlockId: string;
	endOffset: number;
};
