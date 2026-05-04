import { XmlNode } from "./parser";

export const HYPERLINK_RELATIONSHIP_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

export function mintRelationshipId(relationshipsRoot: XmlNode): string {
	let highest = 0;
	for (const child of relationshipsRoot.children) {
		if (child.tag !== "Relationship") continue;
		const id = child.getAttribute("Id");
		if (!id) continue;
		const match = id.match(/^rId(\d+)$/);
		if (match?.[1]) {
			const number = Number(match[1]);
			if (number > highest) highest = number;
		}
	}
	return `rId${highest + 1}`;
}

export function addHyperlinkRelationship(
	relationshipsRoot: XmlNode,
	url: string,
): string {
	const id = mintRelationshipId(relationshipsRoot);
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
