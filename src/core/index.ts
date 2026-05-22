export {
	type BlockReference,
	type CommentReference,
	type DocView,
	enrichImageHashes,
	findBlockById,
	flattenParagraphs,
	type HyperlinkReference,
	type ImageReference,
	openDocView,
	paragraphText,
	paragraphTextAccepted,
	paragraphTextBaseline,
	saveDocView,
	type TrackedChangeReference,
} from "./ast";
export type {
	Block,
	BreakRun,
	ChartRun,
	Comment,
	CommentAnchor,
	Doc,
	DocProperties,
	EquationRun,
	Footnote,
	FootnoteRefRun,
	Hyperlink,
	ImageRun,
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
	type BlockTarget,
	type Locator,
	LocatorParseError,
	LocatorResolveError,
	locatorToBlockTarget,
	parseLocator,
	resolveBlock,
} from "./locators";
export { PkgError } from "./package";
export { nextRelationshipId } from "./package/parts";
export { XmlNode } from "./parser";
export {
	addHyperlinkRelationship,
	HYPERLINK_RELATIONSHIP_TYPE,
} from "./relationships";
export {
	applyColumns,
	applySectionType,
	isSectionType,
	isTrailingSectPr,
	readSectionProperties,
	removeInlineSectPr,
	type SectionProperties,
	SentinelSectionParagraph,
	wrapSectPrChange,
} from "./sections";
export {
	convertTextToDelText,
	createRevisionAllocator,
	isTrackChangesEnabled,
	type RevisionAllocator,
	resolveAuthor,
	resolveDate,
	type TrackedMeta,
} from "./track-changes";
export { Del, Ins, markParagraphMarkAs } from "./track-changes/emit";
