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
