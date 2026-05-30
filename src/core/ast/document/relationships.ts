import { XmlNode } from "../../parser";
import type { ContentTypesView } from "./content-types";
import type { Pkg } from "./package";

const RELATIONSHIPS_PART_NAME = "word/_rels/document.xml.rels";

const HYPERLINK_RELATIONSHIP_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

export const IMAGE_RELATIONSHIP_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

const EMPTY_RELATIONSHIPS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;

type RelationshipInfo = {
	id: string;
	type: string;
	target: string;
	targetMode?: string;
};

export class RelationshipsView {
	tree: XmlNode[];
	/** rId → media part metadata. Populated by `index(contentTypes)` during
	 * `Document.open` (needs the content-types view to resolve each image rel's
	 * content type), kept in sync by `add`/`setTarget`/`remove`. */
	imagesByRelationshipId: Map<
		string,
		{ partName: string; contentType: string }
	> = new Map();
	/** rId → hyperlink URL. Populated by `index(contentTypes)`, kept in sync by
	 * `addHyperlink`/`setTarget`/`remove`. */
	hyperlinksByRelationshipId: Map<string, { url: string }> = new Map();

	constructor(tree: XmlNode[]) {
		this.tree = tree;
	}

	/** Load this view from a package; missing part becomes an empty tree. */
	static async fromPackage(pkg: Pkg): Promise<RelationshipsView> {
		return new RelationshipsView(await pkg.ensurePart(RELATIONSHIPS_PART_NAME));
	}

	/** Parse a view from raw XML; missing input becomes an empty tree. */
	static fromXml(xml?: string): RelationshipsView {
		return new RelationshipsView(XmlNode.parse(xml ?? EMPTY_RELATIONSHIPS_XML));
	}

	/** Serialize this view's tree into the package's `word/_rels/document.xml.rels`. */
	writeTo(pkg: Pkg): void {
		pkg.writeText(RELATIONSHIPS_PART_NAME, XmlNode.serialize(this.tree));
	}

	/** Build the rId → media / rId → hyperlink lookup maps from the tree. Image
	 * rels resolve their `partName`/`contentType` via the content-types view
	 * (hence the dependency); hyperlink rels carry their URL inline. Called once
	 * during `Document.open`/`fromXml`, before the body walk that consumes the
	 * maps to resolve `<a:blip r:embed>` and `<w:hyperlink r:id>`. */
	index(contentTypes: ContentTypesView): void {
		const relationships = XmlNode.findRoot(this.tree, "Relationships");
		if (!relationships) return;
		for (const child of relationships.children) {
			if (child.tag !== "Relationship") continue;
			const type = child.getAttribute("Type");
			const relationshipId = child.getAttribute("Id");
			const target = child.getAttribute("Target");
			if (!relationshipId || !target) continue;
			if (type === IMAGE_RELATIONSHIP_TYPE) {
				const partName = target.startsWith("/")
					? target.slice(1)
					: `word/${target}`;
				this.imagesByRelationshipId.set(relationshipId, {
					partName,
					contentType: contentTypes.lookupContentType(partName),
				});
				continue;
			}
			if (type === HYPERLINK_RELATIONSHIP_TYPE) {
				this.hyperlinksByRelationshipId.set(relationshipId, { url: target });
			}
		}
	}

	list(): RelationshipInfo[] {
		const relationships = XmlNode.findRoot(this.tree, "Relationships");
		if (!relationships) return [];
		const out: RelationshipInfo[] = [];
		for (const child of relationships.children) {
			if (child.tag !== "Relationship") continue;
			const id = child.getAttribute("Id");
			const type = child.getAttribute("Type");
			const target = child.getAttribute("Target");
			if (!id || !type || !target) continue;
			const targetMode = child.getAttribute("TargetMode");
			out.push(
				targetMode ? { id, type, target, targetMode } : { id, type, target },
			);
		}
		return out;
	}

	findByRid(rId: string): XmlNode | undefined {
		const relationships = XmlNode.findRoot(this.tree, "Relationships");
		if (!relationships) return undefined;
		return relationships.children.find(
			(child) =>
				child.tag === "Relationship" && child.getAttribute("Id") === rId,
		);
	}

	findByTarget(target: string): XmlNode | undefined {
		const relationships = XmlNode.findRoot(this.tree, "Relationships");
		if (!relationships) return undefined;
		return relationships.children.find(
			(child) =>
				child.tag === "Relationship" && child.getAttribute("Target") === target,
		);
	}

	/** True if any relationship points at `target` (a relative part path like
	 * `media/image1.png`). Safety gate before deleting a media part on a
	 * format-change replace: identical images are often deduped to one part
	 * shared by several rIds, so we must not delete a part another relationship
	 * still targets. */
	hasTarget(target: string): boolean {
		return this.findByTarget(target) !== undefined;
	}

	/** True if any element in `documentTree` references `rId` through ANY
	 * attribute. Relationship references hide in many attributes we don't model
	 * — `r:embed`/`r:link` (drawings), `r:id` (VML, OLE, `<w:background>`),
	 * `r:dm`/`r:lo`/`r:qs`/`r:cs` (charts) — so we scan every attribute value
	 * rather than a known shortlist. The safety gate before pruning a part:
	 * never delete a relationship/part still referenced by content we can't
	 * parse, or we'd dangle the rId and corrupt the document. */
	isReferenced(rId: string, documentTree: XmlNode[]): boolean {
		function walk(node: XmlNode): boolean {
			for (const value of Object.values(node.attributes)) {
				if (value === rId) return true;
			}
			return node.children.some(walk);
		}
		return documentTree.some(walk);
	}

	/** Mint the next free `rIdN` based on the highest existing numeric suffix. */
	nextId(): string {
		const relationships = XmlNode.findRoot(this.tree, "Relationships");
		if (!relationships) return "rId1";
		let highest = 0;
		for (const child of relationships.children) {
			if (child.tag !== "Relationship") continue;
			const id = child.getAttribute("Id");
			if (!id) continue;
			const match = id.match(/^rId(\d+)$/);
			if (!match) continue;
			const numeric = Number(match[1]);
			if (Number.isFinite(numeric) && numeric > highest) highest = numeric;
		}
		return `rId${highest + 1}`;
	}

	/** Add a `<Relationship Id Type Target [TargetMode]/>` and return its rId.
	 * For OPC-internal parts (styles, numbering, comments, notes), the caller
	 * also calls `ContentTypesView.registerPart(...)`. For media (images), the
	 * caller registers an extension default via `ContentTypesView.registerExtension`. */
	add(type: string, target: string, mode?: "External"): string {
		const relationships = XmlNode.findRoot(this.tree, "Relationships");
		if (!relationships) throw new Error("missing <Relationships> root");
		const id = this.nextId();
		const attributes: Record<string, string> = {
			Id: id,
			Type: type,
			Target: target,
		};
		if (mode) attributes.TargetMode = mode;
		relationships.children.push(new XmlNode("Relationship", attributes));
		return id;
	}

	/** Add a hyperlink relationship and return its rId. Convenience wrapper
	 * around `add(HYPERLINK_RELATIONSHIP_TYPE, url, "External")`. */
	addHyperlink(url: string): string {
		const id = this.add(HYPERLINK_RELATIONSHIP_TYPE, url, "External");
		this.hyperlinksByRelationshipId.set(id, { url });
		return id;
	}

	/** Rewrite the `Target` of the relationship with the given Id (no-op if
	 * absent). Used by `hyperlinks replace` to point an existing rId at a new
	 * URL when the relationship is sole-referenced. */
	setTarget(rId: string, target: string): void {
		const node = this.findByRid(rId);
		if (!node) return;
		node.setAttribute("Target", target);
		// Keep the hyperlink map in sync if this is a hyperlink rel.
		if (this.hyperlinksByRelationshipId.has(rId)) {
			this.hyperlinksByRelationshipId.set(rId, { url: target });
		}
	}

	/** Remove the relationship with the given Id (no-op if absent). */
	remove(rId: string): void {
		const relationships = XmlNode.findRoot(this.tree, "Relationships");
		if (!relationships) return;
		relationships.children = relationships.children.filter(
			(child) =>
				!(child.tag === "Relationship" && child.getAttribute("Id") === rId),
		);
		this.hyperlinksByRelationshipId.delete(rId);
		this.imagesByRelationshipId.delete(rId);
	}

	/** Remove the relationship iff nothing in `documentTree` still references it.
	 * The safety gate around hyperlink/image deletion — see `isReferenced`. */
	removeIfUnreferenced(rId: string, documentTree: XmlNode[]): boolean {
		if (this.isReferenced(rId, documentTree)) return false;
		this.remove(rId);
		return true;
	}
}
