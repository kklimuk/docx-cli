import { PPR_CHILD_ORDER, RPR_CHILD_ORDER } from "../blocks";
import type { XmlNode } from "../parser";
import { SECTPR_CHILD_ORDER } from "../sections";
import { RawError } from "./error";

/** Gate 3: reject a fragment whose ordered-container children are out of their
 *  ECMA-376 sequence. Misordered `<w:pPr>`/`<w:rPr>`/`<w:sectPr>` children are
 *  the #1 cause of Word's "unreadable content / repair" prompt (and of six of
 *  the seven schema-invalid fixtures the validator audit found), so the check
 *  runs on every raw fragment with a precise, named-pair message — the bundled
 *  XSD (gate 6) would catch these too, but only with a generic schema error.
 *  Containers without an order table pass through; the XSD covers them. */
export function checkElementOrder(node: XmlNode): void {
	if (node.isText) return;
	checkLeadingProperty(node);
	checkOrderedContainer(node);
	for (const child of node.children) checkElementOrder(child);
}

/** `w:pPr`-in-`w:p` and friends: the properties element, when present, must be
 *  the FIRST element child of its container. */
const LEADING_PROPERTY: Record<string, string> = {
	"w:p": "w:pPr",
	"w:tbl": "w:tblPr",
	"w:tr": "w:trPr",
	"w:tc": "w:tcPr",
};

function checkLeadingProperty(node: XmlNode): void {
	const property = LEADING_PROPERTY[node.tag];
	if (!property) return;
	const elements = node.children.filter((child) => !child.isText);
	const position = elements.findIndex((child) => child.tag === property);
	if (position > 0) {
		throw new RawError(
			"INVALID_XML",
			`<${property}> must be the first child of <${node.tag}> (ECMA-376 child order)`,
		);
	}
}

function checkOrderedContainer(node: XmlNode): void {
	const order = ORDERED_CONTAINERS.get(node.tag);
	if (!order) return;
	let highestRank = -1;
	let highestTag = "";
	for (const child of node.children) {
		if (child.isText) continue;
		const rank = order.indexOf(child.tag);
		if (rank === -1) continue;
		if (rank < highestRank) {
			throw new RawError(
				"INVALID_XML",
				`<${child.tag}> must precede <${highestTag}> in <${node.tag}> (ECMA-376 child order)`,
			);
		}
		highestRank = rank;
		highestTag = child.tag;
	}
}

const ORDERED_CONTAINERS = new Map<string, readonly string[]>([
	["w:pPr", PPR_CHILD_ORDER],
	["w:rPr", RPR_CHILD_ORDER],
	["w:sectPr", SECTPR_CHILD_ORDER],
]);
