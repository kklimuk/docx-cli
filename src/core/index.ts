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
export {
	FORMAT_TO_NUMFMT,
	type ListFormat,
} from "./ast/document/numbering";
export { PkgError } from "./ast/document/package";
export type {
	Block,
	BreakRun,
	ChartRun,
	Comment,
	CommentAnchor,
	ContentControl,
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
export { ListOperationError, Lists } from "./lists";
export { literalParagraphs } from "./literal-text";
export {
	type BlockRangeReference,
	type BlockTarget,
	type CellReference,
	CellTargetError,
	type CellTargetErrorCode,
	describeForms,
	isCellScopedLocator,
	isMarginalLocator,
	isRelationshipLocator,
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
	resolveCellParagraphReference,
	resolveCellReference,
} from "./locators";
export {
	enumerateMarginalRefs,
	findMarginalRef,
	type MarginalField,
	type MarginalKind,
	type MarginalRef,
	type MarginalSpec,
	Marginals,
	type MarginalType,
} from "./marginals";
export {
	isMarginalType,
	MARGINAL_TYPES,
	marginalConfig,
} from "./marginals/config";
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
