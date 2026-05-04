export { type Locator, LocatorParseError, parseLocator } from "./parse";
export {
	type BlockTarget,
	LocatorResolveError,
	locatorToBlockTarget,
	resolveBlock,
	resolveComment,
	resolveImage,
} from "./resolve";
