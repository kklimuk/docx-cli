import XMLBuilder from "fast-xml-builder";
import { XMLParser } from "fast-xml-parser";

export class XmlNode {
	tag: string;
	attributes: XmlAttributes;
	children: XmlNode[];
	text: string | undefined;

	constructor(
		tag: string,
		attributes: XmlAttributes = {},
		children: XmlNode[] = [],
	) {
		this.tag = tag;
		this.attributes = attributes;
		this.children = children;
	}

	static element(
		tag: string,
		attributes: XmlAttributes = {},
		children: XmlNode[] = [],
	): XmlNode {
		return new XmlNode(tag, attributes, children);
	}

	static textNode(value: string): XmlNode {
		const node = new XmlNode("#text");
		node.text = value;
		return node;
	}

	static parse(xml: string): XmlNode[] {
		const raw = new XMLParser(PARSE_OPTIONS).parse(xml);
		if (!Array.isArray(raw)) return [];
		const tree: XmlNode[] = [];
		for (const item of raw) {
			tree.push(XmlNode.fromObject(item));
		}
		return tree;
	}

	static serialize(tree: XmlNode[]): string {
		const builder = new XMLBuilder(BUILD_OPTIONS);
		const flat: XmlNode[] = [];
		flattenFragments(tree, flat);
		const pojo: Record<string, unknown>[] = [];
		for (const node of flat) {
			pojo.push(node.toObject());
		}
		return builder.build(pojo) as string;
	}

	static serializeWithDeclaration(tree: XmlNode[]): string {
		const withoutDeclaration = tree.filter((node) => node.tag !== "?xml");
		const declaration =
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
		return declaration + XmlNode.serialize(withoutDeclaration);
	}

	static fromObject(pojo: Record<string, unknown>): XmlNode {
		let tagKey: string | undefined;
		for (const key of Object.keys(pojo)) {
			if (key === ":@") continue;
			tagKey = key;
			break;
		}
		if (!tagKey) return new XmlNode("#text");

		const node = new XmlNode(tagKey);

		const rawAttributes = pojo[":@"] as Record<string, string> | undefined;
		if (rawAttributes) {
			for (const [key, value] of Object.entries(rawAttributes)) {
				const stripped = key.startsWith("@_") ? key.slice(2) : key;
				node.attributes[stripped] = value;
			}
		}

		const value = pojo[tagKey];
		if (tagKey === "#text") {
			node.text = typeof value === "string" ? value : "";
			return node;
		}
		if (Array.isArray(value)) {
			for (const child of value) {
				node.children.push(
					XmlNode.fromObject(child as Record<string, unknown>),
				);
			}
		}
		return node;
	}

	static findRoot(tree: XmlNode[], tag: string): XmlNode | undefined {
		for (const node of tree) {
			if (node.tag === tag) return node;
		}
		return undefined;
	}

	get isText(): boolean {
		return this.tag === "#text";
	}

	getAttribute(key: string): string | undefined {
		return this.attributes[key];
	}

	setAttribute(key: string, value: string): void {
		this.attributes[key] = value;
	}

	findChild(tag: string): XmlNode | undefined {
		for (const child of this.children) {
			if (child.tag === tag) return child;
		}
		return undefined;
	}

	findChildren(tag: string): XmlNode[] {
		const matches: XmlNode[] = [];
		for (const child of this.children) {
			if (child.tag === tag) matches.push(child);
		}
		return matches;
	}

	findDescendant(tag: string): XmlNode | undefined {
		for (const child of this.children) {
			if (child.tag === tag) return child;
			const deeper = child.findDescendant(tag);
			if (deeper) return deeper;
		}
		return undefined;
	}

	collectText(): string {
		if (this.isText) return this.text ?? "";
		let out = "";
		for (const child of this.children) {
			out += child.collectText();
		}
		return out;
	}

	toObject(): Record<string, unknown> {
		if (this.isText) {
			return { "#text": this.text ?? "" };
		}
		const result: Record<string, unknown> = {};
		const childObjects: Record<string, unknown>[] = [];
		for (const child of this.children) {
			childObjects.push(child.toObject());
		}
		result[this.tag] = childObjects;

		const attributeKeys = Object.keys(this.attributes);
		if (attributeKeys.length > 0) {
			const rawAttributes: Record<string, string> = {};
			for (const key of attributeKeys) {
				const value = this.attributes[key];
				if (value !== undefined) rawAttributes[`@_${key}`] = value;
			}
			result[":@"] = rawAttributes;
		}
		return result;
	}
}

export type XmlAttributes = Record<string, string>;
export type NullableXmlNode = XmlNode | null;

const PARSE_OPTIONS = {
	ignoreAttributes: false,
	attributeNamePrefix: "@_",
	preserveOrder: true,
	parseAttributeValue: false,
	parseTagValue: false,
	trimValues: false,
	processEntities: true,
	ignoreDeclaration: false,
};

const BUILD_OPTIONS = {
	ignoreAttributes: false,
	attributeNamePrefix: "@_",
	preserveOrder: true,
	suppressEmptyNode: true,
	format: false,
};

const FRAGMENT_TAG = "#fragment";

function flattenFragments(tree: XmlNode[], out: XmlNode[]): void {
	for (const node of tree) {
		if (node.tag === FRAGMENT_TAG) {
			flattenFragments(node.children, out);
			continue;
		}
		out.push(node);
	}
}
