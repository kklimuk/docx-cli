import { type NullableXmlNode, XmlNode } from "../parser";
import { Fragment, type JsxChild } from "./index";

// biome-ignore lint/suspicious/noExplicitAny: components accept arbitrary prop shapes
type Component = (props: any, ...children: JsxChild[]) => NullableXmlNode;

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

type RuntimeProps = { children?: unknown } & Record<string, unknown>;

export function jsx(
	type: Component,
	props: RuntimeProps,
	_key?: unknown,
): XmlNode {
	const { children, ...rest } = props ?? {};
	const childArgs = normalizeChildren(children);
	const result = type(rest, ...childArgs);
	return result ?? new XmlNode("#fragment");
}

export const jsxs = jsx;
export const jsxDEV = jsx;
export { Fragment };

function normalizeChildren(children: unknown): JsxChild[] {
	if (children === undefined) return [];
	if (Array.isArray(children)) return children as JsxChild[];
	return [children as JsxChild];
}
