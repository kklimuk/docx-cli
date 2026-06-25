export {
	type BlockReference,
	Body,
	baselineCatalog,
	type CommentReference,
	Document,
	type EquationReference,
	findBlockById,
	flattenImageRuns,
	flattenParagraphs,
	type HyperlinkReference,
	type ImageReference,
	isBaselineStyle,
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
	Marginal,
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
	CLEARABLE_ATTRS,
	Edit,
	EditError,
	type ParagraphContentSpec,
	type RunFormat,
	resolveClearTags,
} from "./edit";
export { Fonts, type SetDefaultFontResult } from "./fonts";
export {
	Insert,
	InsertError,
	type InsertSpec,
	type TextFormatting,
} from "./insert";
export { literalParagraphs } from "./literal-text";
export {
	type BlockRangeReference,
	type BlockTarget,
	describeForms,
	LOCATOR_FORMS,
	type Locator,
	type LocatorForm,
	type LocatorFormKey,
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
	type MarginalField,
	type MarginalKind,
	type MarginalSpec,
	Marginals,
	type MarginalType,
} from "./marginals";
export { isMarginalType, MARGINAL_TYPES } from "./marginals/config";
export {
	MarkdownImport,
	MarkdownImportError,
	type MarkdownImportErrorCode,
} from "./markdown";
export { XmlNode } from "./parser";
export {
	detectEngine,
	engineByName,
	listAvailable,
	type RenderEngine,
	RenderEngineError,
	type RenderEngineName,
	renderDocxPages,
} from "./render";
export {
	applyPageGeometry,
	getPageContentWidthEmu,
	inheritPageGeometry,
	isSectionType,
	isTrailingSectPr,
	type PageGeometry,
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
