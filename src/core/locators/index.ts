export {
	describeForms,
	LOCATOR_FORMS,
	type LocatorForm,
	type LocatorFormKey,
} from "./forms";
export {
	isCellScopedLocator,
	isMarginalLocator,
	isRelationshipLocator,
	type Locator,
	LocatorParseError,
	parseLocator,
} from "./parse";
export {
	type BlockRangeReference,
	type BlockTarget,
	type CellReference,
	CellTargetError,
	type CellTargetErrorCode,
	LocatorResolveError,
	locatorToBlockTarget,
	parseCellAt,
	parseCellRangeAt,
	parseColumnAt,
	parseRowAt,
	parseTableAt,
	resolveCellParagraphReference,
	resolveCellReference,
} from "./resolve";
