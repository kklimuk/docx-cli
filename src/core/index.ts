export {
	type BlockReference,
	type CommentReference,
	type DocView,
	enrichImageHashes,
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
	resolveImage,
} from "./locators";
export { PkgError } from "./package";
export { XmlNode } from "./parser";
