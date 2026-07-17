import type { Document } from "../ast/document";
import type { XmlNode } from "../parser";
import {
	type MarginalKind,
	type MarginalType,
	marginalConfig,
	marginalPartNameFromTarget,
} from "./config";

/** One resolved header/footer reference: its positional locator id (`hdr0`/
 *  `ftr0`), placement, owning section, the live `<w:sectPr>` it hangs off, and
 *  the part tree it points at. The single source of the positional id →
 *  everything mapping — `read` (→ `Marginal[]`), `headers`/`footers list`,
 *  `raw get --at ftrN`, and `headers`/`footers --at ftrN` all resolve through
 *  here, so the id `list` reports is the id every other surface accepts. */
export type MarginalRef = {
	id: string;
	kind: MarginalKind;
	type: MarginalType;
	sectionId: string;
	sectPr: XmlNode;
	tree: XmlNode[];
};

/** Walk every section's `<w:headerReference>` / `<w:footerReference>` in
 *  document order, resolving each to its part tree. `hdrN`/`ftrN` are positional
 *  per kind (a document-wide marginal — one part referenced from every section —
 *  yields one ref per section, all pointing at the same part). Empty when the
 *  document has no marginals view. */
export function enumerateMarginalRefs(document: Document): MarginalRef[] {
	const refs: MarginalRef[] = [];
	if (!document.marginals) return refs;
	let headerIndex = 0;
	let footerIndex = 0;
	for (const block of document.body.blocks) {
		if (block.type !== "sectionBreak") continue;
		const sectPr = document.body.blockReferences.get(block.id)?.node;
		if (!sectPr) continue;
		for (const child of sectPr.children) {
			const kind = marginalKindForTag(child.tag);
			if (!kind) continue;
			const rId = child.getAttribute("r:id");
			if (!rId) continue;
			const target = document.relationships
				.findByRid(rId)
				?.getAttribute("Target");
			if (!target) continue;
			const partName = marginalPartNameFromTarget(target);
			const tree = document.marginals.partTree(partName);
			if (!tree) continue;
			const type = (child.getAttribute("w:type") ?? "default") as MarginalType;
			const id = `${marginalConfig(kind).locatorPrefix}${
				kind === "header" ? headerIndex++ : footerIndex++
			}`;
			refs.push({ id, kind, type, sectionId: block.id, sectPr, tree });
		}
	}
	return refs;
}

/** Find one marginal reference by its `hdrN`/`ftrN` id, or undefined. */
export function findMarginalRef(
	document: Document,
	marginalId: string,
): MarginalRef | undefined {
	return enumerateMarginalRefs(document).find((ref) => ref.id === marginalId);
}

function marginalKindForTag(tag: string): MarginalKind | undefined {
	if (tag === "w:headerReference") return "header";
	if (tag === "w:footerReference") return "footer";
	return undefined;
}
