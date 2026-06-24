import {
	isMarginalPartName,
	type MarginalKind,
	marginalConfig,
} from "../../marginals/config";
import { XmlNode } from "../../parser";
import type { Pkg } from "./package";
import { RelationshipsView, relsPartNameFor } from "./relationships";

/** One loaded header/footer part: its `<w:hdr>`/`<w:ftr>` tree plus the part's
 *  OWN rels (for a future header-body hyperlink/image, like `NotesView`). */
type MarginalPart = { tree: XmlNode[]; rels?: RelationshipsView };

/** The tree-owning view for ALL header/footer parts (`word/header1.xml`,
 *  `word/footer1.xml`, …). Unlike the single-part views, one `MarginalsView`
 *  owns many parts keyed by part name; the per-section `<w:headerReference>` /
 *  `<w:footerReference>` wiring lives in `word/document.xml`'s `<w:sectPr>`s and
 *  is managed by the `Marginals` lens, not here. The AST reader reads part text
 *  through `partTree`; the lens writes parts through `setPart`. */
export class MarginalsView {
	parts: Map<string, MarginalPart> = new Map();

	/** Load every `word/header*.xml` / `word/footer*.xml` part from the package
	 *  (plus each part's own rels when present). Returns undefined when the doc
	 *  has no marginal parts, matching the optional-view pattern. */
	static async fromPackage(pkg: Pkg): Promise<MarginalsView | undefined> {
		const partNames = pkg.listParts().filter(isMarginalPartName);
		if (partNames.length === 0) return undefined;
		const view = new MarginalsView();
		for (const name of partNames) {
			const tree = await pkg.readPart(name);
			if (!tree) continue;
			const relsTree = await pkg.readPart(relsPartNameFor(name));
			const rels = relsTree
				? new RelationshipsView(relsTree, relsPartNameFor(name))
				: undefined;
			view.parts.set(name, { tree, rels });
		}
		return view;
	}

	/** Serialize every loaded/created part back into the package (plus rels). */
	writeTo(pkg: Pkg): void {
		for (const [name, part] of this.parts) {
			pkg.writeText(name, XmlNode.serialize(part.tree));
			if (part.rels?.hasAny()) part.rels.writeTo(pkg);
		}
	}

	/** The `<w:hdr>`/`<w:ftr>` tree for a part name, or undefined. */
	partTree(name: string): XmlNode[] | undefined {
		return this.parts.get(name)?.tree;
	}

	/** Create or replace a part's tree (its rels, if any, are preserved). */
	setPart(name: string, tree: XmlNode[]): void {
		const existing = this.parts.get(name);
		if (existing) existing.tree = tree;
		else this.parts.set(name, { tree });
	}

	/** Allocate the next free `word/{header|footer}N.xml` across BOTH the parts
	 *  this view already holds AND the package's existing parts (`existingParts`
	 *  = `pkg.listParts()`), so a freshly-minted part never collides with one the
	 *  view didn't load. Indexes are per-kind (`header1`, `footer1`). */
	nextPartName(kind: MarginalKind, existingParts: readonly string[]): string {
		const prefix = marginalConfig(kind).partPrefix;
		const pattern = new RegExp(`^word/${prefix}(\\d+)\\.xml$`);
		let highest = 0;
		const names = new Set<string>([...this.parts.keys(), ...existingParts]);
		for (const name of names) {
			const match = name.match(pattern);
			if (!match) continue;
			const numeric = Number(match[1]);
			if (Number.isFinite(numeric) && numeric > highest) highest = numeric;
		}
		return `word/${prefix}${highest + 1}.xml`;
	}
}
