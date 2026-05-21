import { XmlNode } from "../parser";

export type PartRegistration = {
	partName: string;
	contentType: string;
	relationshipType: string;
	target: string;
};

export function nextRelationshipId(relationships: XmlNode): string {
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

export function registerPart(
	relationshipsTree: XmlNode[],
	contentTypesTree: XmlNode[],
	part: PartRegistration,
): void {
	const relationships = XmlNode.findRoot(relationshipsTree, "Relationships");
	if (relationships) {
		// Match by Type AND Target — a relationship type can be present multiple
		// times for distinct parts (e.g., two `header` rels pointing at header1.xml
		// and header2.xml). Matching by Type alone would false-positive-skip the
		// second registration.
		const alreadyLinked = relationships.children.some(
			(child) =>
				child.tag === "Relationship" &&
				child.getAttribute("Type") === part.relationshipType &&
				child.getAttribute("Target") === part.target,
		);
		if (!alreadyLinked) {
			relationships.children.push(
				new XmlNode("Relationship", {
					Id: nextRelationshipId(relationships),
					Type: part.relationshipType,
					Target: part.target,
				}),
			);
		}
	}

	const types = XmlNode.findRoot(contentTypesTree, "Types");
	if (types) {
		const overrideExists = types.children.some(
			(child) =>
				child.tag === "Override" &&
				child.getAttribute("PartName") === `/${part.partName}`,
		);
		if (!overrideExists) {
			types.children.push(
				new XmlNode("Override", {
					PartName: `/${part.partName}`,
					ContentType: part.contentType,
				}),
			);
		}
	}
}
