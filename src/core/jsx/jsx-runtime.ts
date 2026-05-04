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

export const jsxs = jsx;
export const jsxDEV = jsx;
export { Fragment };
