export {
	type BlockReference,
	Body,
	type CommentReference,
	Document,
	type EquationReference,
	findBlockById,
	flattenImageRuns,
	flattenParagraphs,
	type HyperlinkReference,
	type ImageReference,
	iterateBlocks,
	paragraphText,
	paragraphTextAccepted,
	paragraphTextBaseline,
	type TrackedChangeReference,
} from "./ast";
export { PkgError } from "./ast/document/package";
export type {
	Block,
	BreakRun,
	ChartRun,
	Comment,
	CommentAnchor,
	DocProperties,
	EquationRun,
	Hyperlink,
	ImageRun,
	Note as Footnote,
	NoteRefRun as FootnoteRefRun,
	Paragraph,
	Run,
	SectionBreak,
	SectionType,
	Table,
	TableCell,
	TableRow,
	TableWidth,
	TabRun,
	TextRun,
	TrackedChange,
	TrackedChangeKind,
} from "./ast/types";
export {
	type CommentAnchorSpec,
	Comments,
	CommentsError,
} from "./comments";
export {
	Edit,
	EditError,
	type ParagraphContentSpec,
} from "./edit";
export {
	Insert,
	InsertError,
	type InsertSpec,
	type TextFormatting,
} from "./insert";
export {
	type BlockRangeReference,
	type BlockTarget,
	type Locator,
	LocatorParseError,
	LocatorResolveError,
	locatorToBlockTarget,
	parseCellAt,
	parseCellRangeAt,
	parseColumnAt,
	parseLocator,
	parseRowAt,
	parseTableAt,
} from "./locators";
export {
	MarkdownImport,
	MarkdownImportError,
	type MarkdownImportErrorCode,
} from "./markdown";
export { XmlNode } from "./parser";
export {
	isSectionType,
	isTrailingSectPr,
	readSectionProperties,
	removeInlineSectPr,
	type SectionProperties,
} from "./sections";
export {
	convertTextToDelText,
	type RevisionAllocator,
	resolveAuthor,
	resolveDate,
	TrackChanges,
	type TrackedMeta,
} from "./track-changes";
export { Del, Ins, markParagraphMarkAs } from "./track-changes/emit";
