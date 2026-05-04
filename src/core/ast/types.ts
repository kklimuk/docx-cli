export type Doc = {
	schemaVersion: 1;
	path: string;
	properties: DocProperties;
	blocks: Block[];
	comments: Comment[];
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
};

export type Run = TextRun | ImageRun | BreakRun | TabRun;

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
};

export type ImageRun = {
	type: "image";
	id: string;
	contentType: string;
	hash: string;
	widthEmu?: number;
	heightEmu?: number;
	alt?: string;
	extractedPath?: string;
};

export type BreakRun = {
	type: "break";
	kind: "page" | "line" | "column";
};

export type TabRun = {
	type: "tab";
};

export type TrackedChange = {
	kind: "ins" | "del";
	author: string;
	date: string;
	revisionId: string;
};

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
