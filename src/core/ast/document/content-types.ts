import { XmlNode } from "../../parser";
import type { Pkg } from "./package";

const CONTENT_TYPES_PART_NAME = "[Content_Types].xml";

const EMPTY_CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`;

export class ContentTypesView {
	tree: XmlNode[];

	constructor(tree: XmlNode[]) {
		this.tree = tree;
	}

	/** Load this view from a package; missing part becomes an empty tree. */
	static async fromPackage(pkg: Pkg): Promise<ContentTypesView> {
		return new ContentTypesView(await pkg.ensurePart(CONTENT_TYPES_PART_NAME));
	}

	/** Parse a view from raw XML; missing input becomes an empty tree. */
	static fromXml(xml?: string): ContentTypesView {
		return new ContentTypesView(XmlNode.parse(xml ?? EMPTY_CONTENT_TYPES_XML));
	}

	/** Serialize this view's tree into the package's `[Content_Types].xml`. */
	writeTo(pkg: Pkg): void {
		pkg.writeText(CONTENT_TYPES_PART_NAME, XmlNode.serialize(this.tree));
	}

	/** Resolve a part's content type. Tries per-part `<Override>` first, then
	 * falls back to the `<Default Extension/>` for the part's file extension.
	 * Returns `application/octet-stream` if neither matches — the reader uses
	 * this for image-rel resolution. */
	lookupContentType(partName: string): string {
		const types = XmlNode.findRoot(this.tree, "Types");
		if (!types) return "application/octet-stream";
		for (const child of types.children) {
			if (child.tag !== "Override") continue;
			if (child.getAttribute("PartName") === `/${partName}`) {
				return child.getAttribute("ContentType") ?? "application/octet-stream";
			}
		}
		const extension = partName.split(".").pop()?.toLowerCase() ?? "";
		for (const child of types.children) {
			if (child.tag !== "Default") continue;
			if (child.getAttribute("Extension")?.toLowerCase() === extension) {
				return child.getAttribute("ContentType") ?? "application/octet-stream";
			}
		}
		return "application/octet-stream";
	}

	hasOverride(partName: string): boolean {
		const types = XmlNode.findRoot(this.tree, "Types");
		if (!types) return false;
		return types.children.some(
			(child) =>
				child.tag === "Override" &&
				child.getAttribute("PartName") === `/${partName}`,
		);
	}

	hasDefault(extension: string): boolean {
		const types = XmlNode.findRoot(this.tree, "Types");
		if (!types) return false;
		const lower = extension.toLowerCase();
		return types.children.some(
			(child) =>
				child.tag === "Default" &&
				child.getAttribute("Extension")?.toLowerCase() === lower,
		);
	}

	/** Register a per-part `<Override PartName ContentType/>`. No-op if already
	 * registered. Pairs with `RelationshipsView.add(...)` to allocate a new OPC
	 * part — Document's `ensureXxx()` methods orchestrate both. */
	registerPart(partName: string, contentType: string): void {
		const types = XmlNode.findRoot(this.tree, "Types");
		if (!types) return;
		if (this.hasOverride(partName)) return;
		types.children.push(
			new XmlNode("Override", {
				PartName: `/${partName}`,
				ContentType: contentType,
			}),
		);
	}

	/** Register a `<Default Extension ContentType/>` for parts typed by extension
	 * (media: `image/png`, `image/jpeg`, etc.). No-op if the extension is already
	 * declared. */
	registerExtension(extension: string, contentType: string): void {
		const types = XmlNode.findRoot(this.tree, "Types");
		if (!types) return;
		if (this.hasDefault(extension)) return;
		types.children.push(
			new XmlNode("Default", {
				Extension: extension,
				ContentType: contentType,
			}),
		);
	}
}
