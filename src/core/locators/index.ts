export { type Locator, LocatorParseError, parseLocator } from "./parse";
export {
	type BlockTarget,
	LocatorResolveError,
	locatorToBlockTarget,
	resolveBlock,
	resolveComment,
	resolveHyperlink,
	resolveImage,
} from "./resolve";
