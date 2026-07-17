import type { Document } from "../ast/document";
import type { BlockReference } from "../ast/document/body";
import { XmlNode } from "../parser";
import { checkElementOrder } from "./element-order";
import { RawError } from "./error";
import { ensureFragmentNamespaces, stampRawMarkers } from "./namespaces";
import {
	assertWellFormed,
	type FragmentRoots,
	parseRoots,
	stripInterElementWhitespace,
} from "./parse";
import {
	applyPartAdd,
	applyPartReplace,
	type PreparedPartAdd,
	type PreparedPartReplace,
	preparePartAdd,
	preparePartReplace,
} from "./parts";
import {
	auditFragmentReferences,
	buildReferenceAuditContext,
	type ReferenceAuditContext,
} from "./references";
import {
	applyRelationshipsFragment,
	type PreparedRelationships,
	prepareRelationshipsFragment,
	type RelationshipMode,
	serializeRelationships,
} from "./relationships";

export { RawError, type RawErrorCode } from "./error";
export { assertWellFormed, type FragmentRoots, parseRoots } from "./parse";
export {
	isXmlPartName,
	listPackageParts,
	normalizePartName,
	type PreparedPartAdd,
	type PreparedPartReplace,
} from "./parts";
export {
	isRelationshipFragment,
	type PreparedRelationships,
	type RelationshipMode,
} from "./relationships";

/** The raw-OOXML escape hatch. `docx raw get/insert/replace` let an agent patch
 *  constructs no other verb covers (drop caps, TOC fields, embedded objects,
 *  content controls INSIDE a paragraph, …) — but a raw string is the one input
 *  the rest of the CLI never has to distrust, so every fragment runs a gate
 *  pipeline before it may touch the tree:
 *
 *   1. well-formedness (`XMLValidator`) — our parser NEVER throws; it silently
 *      mangles bad XML (text after a stray `<` is lost), so the gate is separate
 *   2. addressable roots only (`w:p`/`w:tbl`, or `w:sectPr` for a section
 *      replace) — any other block-level element gets no locator and would
 *      silently vanish from `read`, breaking the write-read loop
 *   2.5 inter-element whitespace strip — a pretty-printed fragment parses to
 *      whitespace `#text` children inside `w:p`/`w:tbl`, which Word rejects
 *   3. child-order check against the CT_* order tables we already maintain
 *   4. namespace prefixes — known ones auto-declared on the doc root, unknown
 *      ones rejected (an undeclared prefix is "unreadable content" in Word)
 *   5. reference/id integrity — dangling rIds and missing note/comment ids
 *      reject; drawing/bookmark id collisions are re-minted in the fragment
 *
 *  The full-document XSD check (gate 6) lives in ./validate and runs from the
 *  CLI commit path, where before/after serializations exist. */
export class Raw {
	/** Built once per lens (one document walk) and shared across every
	 *  fragment this instance prepares — a batch re-walking the document per
	 *  entry would also re-mint the same "fresh" drawing/bookmark id twice,
	 *  since fragments aren't in the tree until the splice phase. */
	private referenceAudit?: ReferenceAuditContext;

	constructor(private document: Document) {}

	/** The exact XML of one addressed block — the read half of the
	 *  `get → modify → replace` patch loop. */
	serializeBlock(reference: BlockReference): string {
		return XmlNode.serialize([reference.node]);
	}

	/** Run a user-supplied fragment through gates 1–5 and return splice-ready
	 *  nodes plus non-fatal warnings. Throws {@link RawError} on any reject. */
	prepareFragment(xml: string, allow: FragmentRoots): PreparedFragment {
		assertWellFormed(xml);
		const nodes = parseRoots(xml, allow);
		for (const node of nodes) stripInterElementWhitespace(node);
		for (const node of nodes) checkElementOrder(node);
		this.referenceAudit ??= buildReferenceAuditContext(this.document);
		const warnings = auditFragmentReferences(
			this.document,
			nodes,
			this.referenceAudit,
		);
		ensureFragmentNamespaces(nodes, this.document);
		stampRawMarkers(nodes, this.document);
		return { nodes, warnings };
	}

	/** Splice prepared roots relative to a live target and apply the OOXML
	 *  post-splice rules that belong to the domain, not the CLI: a table cell
	 *  must end with a paragraph (auto-appended — fix, don't hint), adjacent
	 *  sibling tables merge in Word (warn), and replacing a paragraph that
	 *  carries an inline `<w:sectPr>` removes that section boundary (warn).
	 *  Returns the warnings; throws {@link RawError} when the target ref went
	 *  stale (its parent no longer contains it). */
	spliceBlocks(
		target: { node: XmlNode; parent: XmlNode[] },
		action: "after" | "before" | "replace",
		prepared: PreparedFragment,
		options: { cellScoped: boolean },
	): string[] {
		const warnings: string[] = [];
		const index = target.parent.indexOf(target.node);
		if (index === -1) {
			throw new RawError(
				"BLOCK_NOT_FOUND",
				"Block reference is stale (parent does not contain it)",
			);
		}
		if (action === "replace") {
			if (
				target.node.tag === "w:p" &&
				target.node.findChild("w:pPr")?.findChild("w:sectPr")
			) {
				warnings.push(
					"the replaced paragraph carried an inline section break — that section boundary is gone",
				);
			}
			target.parent.splice(index, 1, ...prepared.nodes);
		} else {
			const at = action === "after" ? index + 1 : index;
			target.parent.splice(at, 0, ...prepared.nodes);
		}
		if (options.cellScoped) {
			ensureCellEndsWithParagraph(target.parent, warnings);
		}
		warnOnAdjacentTables(target.parent, prepared.nodes, warnings);
		return warnings;
	}

	/** Re-read the document and prove every spliced root became addressable —
	 *  the write-read-loop invariant, checked BEFORE any save so a failure
	 *  leaves the file untouched. Returns the minted locators in document
	 *  order. */
	proveAddressable(insertedNodes: XmlNode[]): string[] {
		this.document.reread();
		const inserted = new Set(insertedNodes);
		const addressable = new Set<XmlNode>();
		const locators: string[] = [];
		for (const [blockId, reference] of this.document.body.blockReferences) {
			if (!inserted.has(reference.node)) continue;
			addressable.add(reference.node);
			locators.push(blockId);
		}
		for (const node of insertedNodes) {
			if (!addressable.has(node)) {
				throw new RawError(
					"INVALID_XML",
					`Inserted <${node.tag}> did not become addressable after re-reading the document — nothing was written`,
				);
			}
		}
		return locators;
	}

	/** Gate a `<Relationship>` fragment (see ./relationships for the rules). */
	prepareRelationships(
		xml: string,
		mode: RelationshipMode,
	): PreparedRelationships {
		return prepareRelationshipsFragment(this.document, xml, mode);
	}

	/** Splice a prepared relationship fragment into the rels part. */
	applyRelationships(prepared: PreparedRelationships): void {
		applyRelationshipsFragment(this.document, prepared);
	}

	/** One relationship's exact XML (`rId`), or the whole part when omitted.
	 *  Undefined when `rId` names nothing. */
	serializeRelationships(rId?: string): string | undefined {
		return serializeRelationships(this.document, rId);
	}

	/** Gate a new part (see ./parts for the rules). */
	preparePartAdd(
		name: string,
		source: { xml?: string; bytes?: Uint8Array },
		contentType: string | undefined,
	): PreparedPartAdd {
		return preparePartAdd(this.document, name, source, contentType);
	}

	/** Stage a prepared part into the package (+ its content-type Override). */
	applyPartAdd(prepared: PreparedPartAdd): void {
		applyPartAdd(this.document, prepared);
	}

	/** Gate a whole-part replacement. `currentXml` lets a caller that already
	 *  read the part (the `part edit` patch loop) skip the second inflate. */
	preparePartReplace(
		name: string,
		xml: string,
		currentXml?: string,
	): Promise<PreparedPartReplace> {
		return preparePartReplace(this.document, name, xml, currentXml);
	}

	applyPartReplace(prepared: PreparedPartReplace): void {
		applyPartReplace(this.document, prepared);
	}
}

export type PreparedFragment = { nodes: XmlNode[]; warnings: string[] };

/** Word requires a table cell's LAST child to be a paragraph; a raw splice
 *  that leaves a `<w:tbl>` in final position would corrupt the cell, so the
 *  cure is deterministic — append an empty `<w:p>` (fix, don't hint). Applies
 *  only when the splice happened inside a cell. */
function ensureCellEndsWithParagraph(
	parent: XmlNode[],
	warnings: string[],
): void {
	const elements = parent.filter((node) => !node.isText);
	const last = elements[elements.length - 1];
	if (!last || last.tag !== "w:tbl") return;
	parent.push(XmlNode.element("w:p"));
	warnings.push(
		"appended an empty paragraph after the table — Word requires a cell to end with a paragraph",
	);
}

/** Two `<w:tbl>` siblings with nothing between them MERGE into one table when
 *  Word opens the file — legal XML, surprising result. Warn, don't block. */
function warnOnAdjacentTables(
	parent: XmlNode[],
	insertedNodes: XmlNode[],
	warnings: string[],
): void {
	if (!insertedNodes.some((node) => node.tag === "w:tbl")) return;
	const elements = parent.filter((node) => !node.isText);
	for (const node of insertedNodes) {
		if (node.tag !== "w:tbl") continue;
		const index = elements.indexOf(node);
		const neighbors = [elements[index - 1], elements[index + 1]];
		if (
			neighbors.some(
				(sibling) =>
					sibling &&
					sibling.tag === "w:tbl" &&
					!insertedNodes.includes(sibling),
			)
		) {
			warnings.push(
				"the inserted table touches an existing table — adjacent tables merge into one in Word; insert an empty paragraph between them to keep two",
			);
			return;
		}
	}
}
