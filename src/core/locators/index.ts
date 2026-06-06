export {
	describeForms,
	LOCATOR_FORMS,
	type LocatorForm,
	type LocatorFormKey,
} from "./forms";
export { type Locator, LocatorParseError, parseLocator } from "./parse";
export {
	type BlockRangeReference,
	type BlockTarget,
	LocatorResolveError,
	locatorToBlockTarget,
	parseCellAt,
	parseCellRangeAt,
	parseColumnAt,
	parseRowAt,
	parseTableAt,
} from "./resolve";
