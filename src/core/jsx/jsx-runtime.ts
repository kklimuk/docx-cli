import { type NullableXmlNode, XmlNode } from "../parser";
import { Fragment } from "./index";

declare global {
	namespace JSX {
		type Element = XmlNode;
		interface IntrinsicElements {
			[tag: string]: Record<string, unknown>;
		}
		interface ElementChildrenAttribute {
			children: object;
		}
	}
}

export function jsx<P>(
	type: (props: P) => NullableXmlNode,
	props: P,
	_key?: unknown,
): XmlNode {
	const result = type(props);
	return result ?? new XmlNode("#fragment");
}

// The automatic runtime calls `jsxs` for elements with statically-known
// children and `jsxDEV` in development; our builder treats all three
// identically. They delegate to `jsx` as distinct functions (rather than
// `= jsx` aliases) so each is a single unambiguous export.
export function jsxs<P>(
	type: (props: P) => NullableXmlNode,
	props: P,
	key?: unknown,
): XmlNode {
	return jsx(type, props, key);
}

export function jsxDEV<P>(
	type: (props: P) => NullableXmlNode,
	props: P,
	key?: unknown,
): XmlNode {
	return jsx(type, props, key);
}

export { Fragment };
