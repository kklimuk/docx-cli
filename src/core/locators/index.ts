export { type Locator, LocatorParseError, parseLocator } from "./parse";
export {
	type BlockRangeReference,
	type BlockTarget,
	LocatorResolveError,
	locatorToBlockTarget,
	resolveBlock,
	resolveBlockRange,
} from "./resolve";
