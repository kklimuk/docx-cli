export {
	type BlockReference,
	type CommentReference,
	type DocView,
	enrichImageHashes,
	type HyperlinkReference,
	type ImageReference,
	openDocView,
	saveDocView,
} from "./ast";
export type {
	Block,
	BreakRun,
	Comment,
	CommentAnchor,
	Doc,
	DocProperties,
	Hyperlink,
	ImageRun,
	Paragraph,
	Run,
	SectionBreak,
	Table,
	TableCell,
	TableRow,
	TabRun,
	TextRun,
	TrackedChange,
} from "./ast/types";
export {
	type BlockTarget,
	type Locator,
	LocatorParseError,
	LocatorResolveError,
	locatorToBlockTarget,
	parseLocator,
	resolveBlock,
	resolveComment,
	resolveHyperlink,
	resolveImage,
} from "./locators";
export { PkgError } from "./package";
export { XmlNode } from "./parser";
export {
	addHyperlinkRelationship,
	HYPERLINK_RELATIONSHIP_TYPE,
	mintRelationshipId,
} from "./relationships";
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
