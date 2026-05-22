import { nextRelationshipId } from "./package/parts";
import { XmlNode } from "./parser";

export const HYPERLINK_RELATIONSHIP_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

export function addHyperlinkRelationship(
	relationshipsRoot: XmlNode,
	url: string,
): string {
	const id = nextRelationshipId(relationshipsRoot);
	relationshipsRoot.children.push(
		new XmlNode("Relationship", {
			Id: id,
			Type: HYPERLINK_RELATIONSHIP_TYPE,
			Target: url,
			TargetMode: "External",
		}),
	);
	return id;
}

/** Rewrite the `Target` of the relationship with the given Id (no-op if absent). */
export function setRelationshipTarget(
	relationshipsTree: XmlNode[],
	relationshipId: string,
	target: string,
): void {
	const relationships = XmlNode.findRoot(relationshipsTree, "Relationships");
	if (!relationships) return;
	for (const child of relationships.children) {
		if (
			child.tag === "Relationship" &&
			child.getAttribute("Id") === relationshipId
		) {
			child.setAttribute("Target", target);
			return;
		}
	}
}

/** True if any element in the tree references the relationship id through ANY
 * attribute. Relationship references hide in many attributes we don't model —
 * `r:embed`/`r:link` (drawings), `r:id` (VML `<v:imagedata>`, OLE objects,
 * `<w:background>`), `r:dm`/`r:lo`/`r:qs`/`r:cs` (charts) — so we scan every
 * attribute value rather than a known shortlist. This is the safety gate before
 * pruning a part: never delete a relationship/part still referenced by content
 * we can't parse, or we'd dangle the rId and corrupt the document. */
export function isRelationshipReferenced(
	tree: XmlNode[],
	relationshipId: string,
): boolean {
	function walk(node: XmlNode): boolean {
		for (const value of Object.values(node.attributes)) {
			if (value === relationshipId) return true;
		}
		return node.children.some(walk);
	}
	return tree.some(walk);
}

/** True if any relationship points at `target` (a relative part path like
 * `media/image1.png`). The safety gate before deleting a media part on a
 * format-change replace: identical images are often deduped to one part shared
 * by several rIds, so we must not delete a part another relationship still
 * targets. */
export function hasRelationshipWithTarget(
	relationshipsTree: XmlNode[],
	target: string,
): boolean {
	const relationships = XmlNode.findRoot(relationshipsTree, "Relationships");
	if (!relationships) return false;
	return relationships.children.some(
		(child) =>
			child.tag === "Relationship" && child.getAttribute("Target") === target,
	);
}

/** Remove the relationship with the given Id (no-op if absent). */
export function removeRelationship(
	relationshipsTree: XmlNode[],
	relationshipId: string,
): void {
	const relationships = XmlNode.findRoot(relationshipsTree, "Relationships");
	if (!relationships) return;
	relationships.children = relationships.children.filter(
		(child) =>
			!(
				child.tag === "Relationship" &&
				child.getAttribute("Id") === relationshipId
			),
	);
}
