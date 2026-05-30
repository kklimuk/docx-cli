import { XmlNode } from "../../parser";
import type { DocProperties } from "../types";
import type { Pkg } from "./package";

const CORE_PROPERTIES_PART_NAME = "docProps/core.xml";

/** Read-only view over `docProps/core.xml`. No CLI verb currently mutates the
 * core properties — the view exists to surface them on `Body.properties` for the
 * `--ast` and `info` outputs. */
export class CorePropertiesView {
	tree: XmlNode[];

	constructor(tree: XmlNode[]) {
		this.tree = tree;
	}

	/** Load this view from a package; returns undefined if the part is absent. */
	static async fromPackage(pkg: Pkg): Promise<CorePropertiesView | undefined> {
		const tree = await pkg.readPart(CORE_PROPERTIES_PART_NAME);
		return tree ? new CorePropertiesView(tree) : undefined;
	}

	/** Parse a view from raw XML; returns undefined if the input is absent. */
	static fromXml(xml: string | undefined): CorePropertiesView | undefined {
		return xml ? new CorePropertiesView(XmlNode.parse(xml)) : undefined;
	}

	/** Serialize this view's tree into the package's `docProps/core.xml`. */
	writeTo(pkg: Pkg): void {
		pkg.writeText(CORE_PROPERTIES_PART_NAME, XmlNode.serialize(this.tree));
	}

	get(key: "title" | "author" | "created" | "modified"): string | undefined {
		const root = XmlNode.findRoot(this.tree, "cp:coreProperties");
		if (!root) return undefined;
		const child = root.findChild(coreTagFor(key));
		return child?.collectText();
	}

	/** Snapshot the four well-known properties (title / author / created /
	 * modified) into a `DocProperties` object — the shape `Body.properties`
	 * exposes for `--ast` consumers. Omits keys the part doesn't carry. */
	snapshot(): DocProperties {
		const out: DocProperties = {};
		const title = this.get("title");
		if (title !== undefined) out.title = title;
		const author = this.get("author");
		if (author !== undefined) out.author = author;
		const created = this.get("created");
		if (created !== undefined) out.created = created;
		const modified = this.get("modified");
		if (modified !== undefined) out.modified = modified;
		return out;
	}
}

function coreTagFor(key: "title" | "author" | "created" | "modified"): string {
	switch (key) {
		case "title":
			return "dc:title";
		case "author":
			return "dc:creator";
		case "created":
			return "dcterms:created";
		case "modified":
			return "dcterms:modified";
	}
}
