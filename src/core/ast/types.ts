export type Doc = {
	schemaVersion: 1;
	path: string;
	properties: DocProperties;
	blocks: Block[];
	comments: Comment[];
	footnotes: Footnote[];
	endnotes: Footnote[];
};

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
	runs: Run[];
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
	rows: TableRow[];
};

export type TableRow = {
	cells: TableCell[];
	/** Tracked row insertion/deletion from <w:trPr><w:ins/> or <w:del/>
	 * (kind "rowIns" / "rowDel"). Present only under track-changes. */
	trackedChange?: TrackedChange;
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
	| FootnoteRefRun
	| ChartRun;

export type TextRun = {
	type: "text";
	text: string;
	color?: string;
	highlight?: string;
	bold?: boolean;
	italic?: boolean;
	underline?: string;
	strike?: boolean;
	font?: string;
	sizeHalfPoints?: number;
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
	trackedChange?: TrackedChange;
};

export type BreakRun = {
	type: "break";
	kind: "page" | "line" | "column";
};

export type TabRun = {
	type: "tab";
};

/** A math equation surfaced as concatenated <m:t> plaintext. The structural
 * OOMath markup (subscripts, fractions, etc.) is collapsed to literal characters,
 * so the rendering is degraded. `display` distinguishes inline equations
 * (<m:oMath>) from block-level display equations (<m:oMathPara>). */
export type EquationRun = {
	type: "equation";
	text: string;
	display: boolean;
};

/** Reference to a footnote or endnote whose body lives in Doc.footnotes or
 * Doc.endnotes. The id matches the footnote/endnote id (`fn1`, `en1`). */
export type FootnoteRefRun = {
	type: "footnoteRef";
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

export type Footnote = {
	id: string;
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
 */
export type TrackedChangeKind =
	| "ins"
	| "del"
	| "moveFrom"
	| "moveTo"
	| "sectPrChange"
	| "rowIns"
	| "rowDel"
	| "cellIns"
	| "cellDel"
	| "tblGridChange"
	| "tblPrChange"
	| "tcPrChange";

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
