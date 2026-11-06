import { XmlNode } from "../../parser";
import type { Pkg } from "./package";
import {
	RelationshipsView,
	relsPartNameFor,
	resolveRelationshipTarget,
} from "./relationships";

/** One chart part: its `<c:chartSpace>` tree plus its OWN rels (where
 * `<c:externalData r:id>` resolves to the embedded or linked workbook). */
export type ChartPart = {
	tree: XmlNode[];
	relationships: RelationshipsView;
};

/** The chart parts (`word/charts/chartN.xml`) the document body references,
 * keyed by part name — the multi-part shape `MarginalsView` set. READ-ONLY for
 * now: the reader parses the cached series out of these trees, and nothing
 * mutates them, so there is deliberately no `writeTo` — an unrelated save
 * never re-serializes a chart part, and the chart, its rels and its workbook
 * come through byte-identical. A future `charts set-data` that edits them adds
 * the write path. */
export class ChartsView {
	constructor(private readonly parts: Map<string, ChartPart>) {}

	/** Load every part a `chart` relationship of `word/document.xml` targets
	 * (transitional or Strict relationship type), with its rels part. */
	static async fromPackage(
		pkg: Pkg,
		documentRelationships: RelationshipsView,
	): Promise<ChartsView> {
		const parts = new Map<string, ChartPart>();
		for (const relationship of documentRelationships.list()) {
			if (!isChartRelationshipType(relationship.type)) continue;
			if (relationship.targetMode === "External") continue;
			const partName = resolveRelationshipTarget(
				"word/document.xml",
				relationship.target,
			);
			if (parts.has(partName)) continue;
			const tree = await pkg.readPart(partName);
			if (!tree) continue;
			const relsPartName = relsPartNameFor(partName);
			parts.set(partName, {
				tree,
				relationships: await RelationshipsView.fromPackage(pkg, relsPartName),
			});
		}
		return new ChartsView(parts);
	}

	/** Build a view from raw chart XML keyed by part name — the synthetic
	 * `Document.fromXml` path for reader tests. */
	static fromXml(
		parts: Record<string, { xml: string; relsXml?: string }> = {},
	): ChartsView {
		const map = new Map<string, ChartPart>();
		for (const [partName, { xml, relsXml }] of Object.entries(parts)) {
			map.set(partName, {
				tree: XmlNode.parse(xml),
				relationships: RelationshipsView.fromXml(
					relsXml,
					relsPartNameFor(partName),
				),
			});
		}
		return new ChartsView(map);
	}

	get(partName: string): ChartPart | undefined {
		return this.parts.get(partName);
	}
}

/** `…/officeDocument/2006/relationships/chart` (transitional) or
 * `http://purl.oclc.org/ooxml/officeDocument/relationships/chart` (Strict). */
function isChartRelationshipType(type: string): boolean {
	return type.endsWith("/relationships/chart");
}
