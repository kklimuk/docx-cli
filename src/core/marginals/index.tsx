import type { Document } from "../ast/document";
import { w } from "../jsx";
import type { XmlNode } from "../parser";
import { insertSectPrChildInOrder, wrapSectPrChange } from "../sections";
import { TrackChanges } from "../track-changes";
import {
	type MarginalConfig,
	type MarginalKind,
	type MarginalType,
	marginalConfig,
	marginalPartNameFromTarget,
} from "./config";

export type {
	MarginalConfig,
	MarginalKind,
	MarginalType,
} from "./config";

/** A header/footer field: rides one `<w:fldSimple>` (or a small run sequence for
 *  "Page X of Y"). `page` → PAGE; `date` → DATE (optional `\@` format); `styleRef`
 *  → STYLEREF (running head); `filename`/`title`/`author` → the doc-property fields. */
export type MarginalField =
	| { type: "page"; ofPages?: boolean }
	| { type: "date"; format?: string }
	| { type: "styleRef"; style: string }
	| { type: "filename" }
	| { type: "title" }
	| { type: "author" };

/** What to put in a header/footer. ONE primary source (`text` or `field`), except
 *  `text` + `field` together = two-zone (text left, field right via a tab at the
 *  content edge). `align` positions the single-zone case (ignored for two-zone). */
export type MarginalSpec = {
	text?: string;
	align?: "left" | "center" | "right";
	field?: MarginalField;
};

/** Cross-cutting lens for headers/footers — constructed at the call site like
 *  `Images`/`TrackChanges`. Reaches through `Document` into the `MarginalsView`
 *  (the part trees), the relationships + content-types views (part registration),
 *  the settings view (the even/odd toggle), and the live `<w:sectPr>` nodes (the
 *  per-section references). The reference wiring rides the existing
 *  `<w:sectPrChange>` machinery under `--track`. */
export class Marginals {
	constructor(private document: Document) {}

	/** Set (create or replace) a header/footer of `type` on each of `sectPrs`,
	 *  sharing ONE part across all of them (one rId referenced from each section,
	 *  so "document-wide" is one part, many references). Returns the part name +
	 *  rId + how many sections now reference it. Tracking wraps a section in
	 *  `<w:sectPrChange>` only when ITS reference actually changes (a pure content
	 *  replace of an existing reference is not section-tracked in v1). */
	set(
		sectPrs: XmlNode[],
		kind: MarginalKind,
		type: MarginalType,
		spec: MarginalSpec,
		opts: { track?: boolean; authorFlag?: string } = {},
	): {
		partName: string;
		rId: string;
		sections: number;
		referencesChanged: number;
	} {
		const config = marginalConfig(kind);
		const view = this.document.ensureMarginals();

		// Reuse an existing (kind,type) reference's part IN PLACE only when no
		// section OUTSIDE the target set references it — otherwise overwriting its
		// shared body would silently rewrite those other sections (the `--at sN`
		// "one section" contract). When the part IS shared with a non-target
		// section, COPY-ON-WRITE: mint a fresh part for the targeted sections and
		// repoint only their references, leaving the others on the shared part.
		const existingId = this.findExistingReferenceId(sectPrs, config, type);
		const existingPart = existingId
			? this.resolvePartName(existingId)
			: undefined;
		const canReuseInPlace =
			existingId !== undefined &&
			existingPart !== undefined &&
			!this.referencedOutsideTargets(existingId, config, sectPrs);
		let rId: string;
		let partName: string;
		if (canReuseInPlace) {
			rId = existingId as string;
			partName = existingPart as string;
		} else {
			partName = view.nextPartName(kind, this.document.pkg.listParts());
			rId = this.document.relationships.add(
				config.relationshipType,
				relationshipTarget(partName),
			);
			this.document.contentTypes.registerPart(partName, config.contentType);
		}

		// Build + write the shared part body (geometry from the first section).
		const firstSectPr = sectPrs[0];
		const contentWidth = firstSectPr ? contentWidthTwips(firstSectPr) : 9360;
		const paragraph = buildContentParagraph(spec, contentWidth);
		view.setPart(partName, [buildMarginalRoot(config, paragraph)]);

		let referencesChanged = 0;
		for (const sectPr of sectPrs) {
			const reference = findReference(sectPr, config.referenceTag, type);
			const referenceChanged =
				!reference || reference.getAttribute("r:id") !== rId;
			if (referenceChanged) referencesChanged++;
			if (referenceChanged && opts.track) {
				wrapSectPrChange(
					sectPr,
					new TrackChanges(this.document).mintMeta(opts.authorFlag),
				);
			}
			if (reference) reference.setAttribute("r:id", rId);
			else insertSectPrChildInOrder(sectPr, buildReference(config, type, rId));
			// First-page-different needs <w:titlePg/> on the section, or Word
			// ignores a `first`-type reference entirely.
			if (type === "first") ensureTitlePg(sectPr);
		}

		// Even/odd headers are gated by a DOCUMENT-level toggle in settings.xml.
		if (type === "even") {
			this.document.ensureSettings().ensureEvenAndOddHeaders();
		}

		return { partName, rId, sections: sectPrs.length, referencesChanged };
	}

	/** Remove a header/footer of `type` from each of `sectPrs` (just the
	 *  `<w:…Reference>`; the part + relationship are left as harmless orphans per
	 *  the unreferenced-part invariant). Returns how many references were removed.
	 *  Tracking snapshots each section so reject restores the reference. */
	clear(
		sectPrs: XmlNode[],
		kind: MarginalKind,
		type: MarginalType,
		opts: { track?: boolean; authorFlag?: string } = {},
	): { removed: number } {
		const config = marginalConfig(kind);
		let removed = 0;
		const affected: XmlNode[] = [];
		for (const sectPr of sectPrs) {
			const reference = findReference(sectPr, config.referenceTag, type);
			if (!reference) continue;
			if (opts.track) {
				wrapSectPrChange(
					sectPr,
					new TrackChanges(this.document).mintMeta(opts.authorFlag),
				);
			}
			const index = sectPr.children.indexOf(reference);
			if (index !== -1) {
				sectPr.children.splice(index, 1);
				removed++;
				affected.push(sectPr);
			}
		}
		// `first` is gated by a per-section `<w:titlePg/>`; once a section has
		// neither a first header NOR a first footer, that toggle is an orphan that
		// wrongly suppresses page-1 DEFAULT content — drop it so clearing the
		// first-page header restores the default on page 1. (The sectPrChange
		// snapshot above already captured titlePg, so reject still restores it.)
		if (type === "first") {
			for (const sectPr of affected) {
				if (!this.sectionHasMarginalType(sectPr, "first"))
					removeTitlePg(sectPr);
			}
		}
		// `even` is gated by the DOCUMENT-level `<w:evenAndOddHeaders/>`; once NO
		// section anywhere has an even header or footer, a stale toggle leaves even
		// pages blank instead of inheriting the default — remove it.
		if (type === "even" && removed > 0 && !this.anySectionHasType("even")) {
			this.document.ensureSettings().removeEvenAndOddHeaders();
		}
		return { removed };
	}

	/** Every `<w:sectPr>` in the document body (inline + trailing), via the AST
	 *  section blocks. Used to detect whether a part is shared beyond the targets. */
	private allSectPrs(): XmlNode[] {
		const result: XmlNode[] = [];
		for (const block of this.document.body.blocks) {
			if (block.type !== "sectionBreak") continue;
			const node = this.document.body.blockReferences.get(block.id)?.node;
			if (node) result.push(node);
		}
		return result;
	}

	/** True when `sectPr` carries ANY marginal reference (header OR footer) of
	 *  `type`. Used to decide whether a per-section toggle (`<w:titlePg/>`) is now
	 *  an orphan after a clear. */
	private sectionHasMarginalType(sectPr: XmlNode, type: MarginalType): boolean {
		const tags: string[] = [
			marginalConfig("header").referenceTag,
			marginalConfig("footer").referenceTag,
		];
		return sectPr.children.some(
			(child) =>
				tags.includes(child.tag) &&
				(child.getAttribute("w:type") ?? "default") === type,
		);
	}

	/** True when ANY section in the document has a marginal reference (header or
	 *  footer) of `type` — for the document-level even/odd toggle. */
	private anySectionHasType(type: MarginalType): boolean {
		return this.allSectPrs().some((sectPr) =>
			this.sectionHasMarginalType(sectPr, type),
		);
	}

	/** True when a section OUTSIDE `targets` carries a `<w:{kind}Reference>` (any
	 *  type) pointing at `rId` — i.e. the part backing `rId` is shared, so
	 *  overwriting it in place would corrupt that other section. */
	private referencedOutsideTargets(
		rId: string,
		config: MarginalConfig,
		targets: XmlNode[],
	): boolean {
		const targetSet = new Set(targets);
		for (const sectPr of this.allSectPrs()) {
			if (targetSet.has(sectPr)) continue;
			const shares = sectPr.children.some(
				(child) =>
					child.tag === config.referenceTag &&
					child.getAttribute("r:id") === rId,
			);
			if (shares) return true;
		}
		return false;
	}

	private findExistingReferenceId(
		sectPrs: XmlNode[],
		config: MarginalConfig,
		type: MarginalType,
	): string | undefined {
		for (const sectPr of sectPrs) {
			const reference = findReference(sectPr, config.referenceTag, type);
			const id = reference?.getAttribute("r:id");
			if (id) return id;
		}
		return undefined;
	}

	private resolvePartName(rId: string): string | undefined {
		const relationship = this.document.relationships.findByRid(rId);
		const target = relationship?.getAttribute("Target");
		return target ? marginalPartNameFromTarget(target) : undefined;
	}
}

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function buildMarginalRoot(
	config: MarginalConfig,
	paragraph: XmlNode,
): XmlNode {
	const Root = config.kind === "header" ? w.hdr : w.ftr;
	// `r:` is declared for a future header-body hyperlink/image, exactly as the
	// notes parts declare it — an undeclared prefix is malformed XML.
	return <Root {...{ "xmlns:w": NS_W, "xmlns:r": NS_R }}>{paragraph}</Root>;
}

function buildReference(
	config: MarginalConfig,
	type: MarginalType,
	rId: string,
): XmlNode {
	const Reference =
		config.kind === "header" ? w.headerReference : w.footerReference;
	return <Reference w-type={type} r-id={rId} />;
}

/** Build the single content paragraph. Three shapes: text+field → two-zone (text
 *  left, field right at a content-edge tab); field only → aligned field; text only
 *  → aligned text. */
function buildContentParagraph(
	spec: MarginalSpec,
	contentWidth: number,
): XmlNode {
	const hasText = spec.text !== undefined && spec.text !== "";
	const field = spec.field;
	if (hasText && field) {
		return (
			<w.p>
				<w.pPr>
					<w.tabs>
						<w.tab w-val="right" w-pos={String(contentWidth)} />
					</w.tabs>
				</w.pPr>
				{textRun(spec.text as string)}
				<w.r>
					<w.tab />
				</w.r>
				{fieldRuns(field)}
			</w.p>
		);
	}
	if (field) {
		const alignment = alignChild(spec.align ?? "center");
		return (
			<w.p>
				{alignment ? <w.pPr>{alignment}</w.pPr> : null}
				{fieldRuns(field)}
			</w.p>
		);
	}
	const alignment = alignChild(spec.align ?? "left");
	return (
		<w.p>
			{alignment ? <w.pPr>{alignment}</w.pPr> : null}
			{hasText ? textRun(spec.text as string) : null}
		</w.p>
	);
}

function alignChild(align: "left" | "center" | "right"): XmlNode | null {
	if (align === "left") return null;
	return <w.jc w-val={align} />;
}

function textRun(text: string): XmlNode {
	return (
		<w.r>
			<w.t {...{ "xml:space": "preserve" }}>{text}</w.t>
		</w.r>
	);
}

function fieldRuns(field: MarginalField): XmlNode[] {
	switch (field.type) {
		case "page":
			if (field.ofPages) {
				return [
					textRun("Page "),
					fieldSimple(" PAGE ", "1"),
					textRun(" of "),
					fieldSimple(" NUMPAGES ", "1"),
				];
			}
			return [fieldSimple(" PAGE ", "1")];
		case "date":
			return [
				fieldSimple(
					field.format ? ` DATE \\@ "${field.format}" ` : " DATE ",
					"",
				),
			];
		case "styleRef":
			return [fieldSimple(` STYLEREF "${field.style}" `, "")];
		case "filename":
			return [fieldSimple(" FILENAME \\p ", "")];
		case "title":
			return [fieldSimple(" TITLE ", "")];
		case "author":
			return [fieldSimple(" AUTHOR ", "")];
	}
}

/** A `<w:fldSimple>` with a cached placeholder run (Word/LibreOffice recompute on
 *  open; the placeholder is the pre-recompute fallback). */
function fieldSimple(instr: string, cached: string): XmlNode {
	return (
		<w.fldSimple w-instr={instr}>
			<w.r>
				<w.t {...{ "xml:space": "preserve" }}>{cached}</w.t>
			</w.r>
		</w.fldSimple>
	);
}

function findReference(
	sectPr: XmlNode,
	tag: string,
	type: MarginalType,
): XmlNode | undefined {
	return sectPr.children.find(
		(child) =>
			child.tag === tag && (child.getAttribute("w:type") ?? "default") === type,
	);
}

function ensureTitlePg(sectPr: XmlNode): void {
	if (sectPr.findChild("w:titlePg")) return;
	insertSectPrChildInOrder(sectPr, <w.titlePg />);
}

function removeTitlePg(sectPr: XmlNode): void {
	const index = sectPr.children.findIndex((child) => child.tag === "w:titlePg");
	if (index !== -1) sectPr.children.splice(index, 1);
}

/** A section's content width in twips (page width − L/R margins), for the
 *  two-zone right tab. Falls back to US-Letter-with-1″ when the sectPr omits
 *  geometry (it inherits — the document section carries it). */
function contentWidthTwips(sectPr: XmlNode): number {
	const width =
		intAttr(sectPr.findChild("w:pgSz")?.getAttribute("w:w")) ?? 12240;
	const pgMar = sectPr.findChild("w:pgMar");
	const left = intAttr(pgMar?.getAttribute("w:left")) ?? 1440;
	const right = intAttr(pgMar?.getAttribute("w:right")) ?? 1440;
	const contentWidth = width - left - right;
	return contentWidth > 0 ? contentWidth : 9360;
}

function intAttr(raw: string | undefined): number | undefined {
	if (raw === undefined) return undefined;
	const parsed = Number.parseInt(raw, 10);
	return Number.isFinite(parsed) ? parsed : undefined;
}

/** The relationship `Target` for a part — relative to `word/` (the
 *  `word/_rels/document.xml.rels` base), so `word/header1.xml` → `header1.xml`. */
function relationshipTarget(partName: string): string {
	return partName.startsWith("word/")
		? partName.slice("word/".length)
		: partName;
}
