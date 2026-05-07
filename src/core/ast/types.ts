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
	list?: { level: number; numId: string };
	runs: Run[];
};

export type Table = {
	id: string;
	type: "table";
	rows: TableRow[];
};

export type TableRow = {
	cells: TableCell[];
};

export type TableCell = {
	blocks: Block[];
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
 */
export type TrackedChangeKind =
	| "ins"
	| "del"
	| "moveFrom"
	| "moveTo"
	| "sectPrChange";

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
