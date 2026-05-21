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
