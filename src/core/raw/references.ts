import type { Document } from "../ast/document";
import { isDrawingIdCarrier } from "../image/drawing";
import { prefixOf } from "../mc";
import type { XmlNode } from "../parser";
import { RawError } from "./error";

/** Gate 5: audit every id the fragment references or defines against the
 *  document. Split three ways by what the tool can honestly do:
 *   - a dangling REFERENCE (rId, footnote/endnote/comment id) → ERROR — it
 *     corrupts the file or points at nothing, and we can't invent the target
 *   - a colliding DEFINITION (`wp:docPr`/`pic:cNvPr` id, bookmark id) →
 *     auto-remint inside the fragment — deterministic, invisible, fix-don't-hint
 *   - a missing STYLE/NUMBERING reference → WARNING — Word falls back cleanly
 *  Returns the warnings; throws {@link RawError} on the errors.
 *
 *  The remint passes draw from a {@link ReferenceAuditContext} the caller
 *  builds ONCE (one document walk) and reuses across fragments — a batch that
 *  audited each entry against the tree alone would both re-walk the document
 *  per entry and mint the same "fresh" id twice (fragments aren't in the tree
 *  until the splice phase). */
export function auditFragmentReferences(
	document: Document,
	nodes: XmlNode[],
	context: ReferenceAuditContext,
): string[] {
	const warnings: string[] = [];
	checkRelationshipReferences(document, nodes);
	checkNoteAndCommentReferences(document, nodes);
	remintDrawingIds(nodes, context);
	remintBookmarkIds(nodes, context);
	warnOnUnknownStyles(document, nodes, warnings);
	warnOnUnknownNumbering(document, nodes, warnings);
	return warnings;
}

/** The document's id landscape for the remint passes: every `wp:docPr`/
 *  `pic:cNvPr` id and bookmark id in use, plus the next free cursor for each.
 *  Minted ids are recorded back into the sets, so one context stays
 *  collision-free across many fragments. */
export type ReferenceAuditContext = {
	drawingIds: Set<string>;
	nextDrawingId: number;
	bookmarkIds: Set<string>;
	nextBookmarkId: number;
};

export function buildReferenceAuditContext(
	document: Document,
): ReferenceAuditContext {
	const drawingIds = new Set<string>();
	const bookmarkIds = new Set<string>();
	for (const root of document.documentTree) {
		walk(root, (element) => {
			if (isDrawingIdCarrier(element.tag)) {
				const id = element.getAttribute("id");
				if (id !== undefined) drawingIds.add(id);
			}
			if (isBookmarkMarker(element.tag)) {
				const id = element.getAttribute("w:id");
				if (id !== undefined) bookmarkIds.add(id);
			}
		});
	}
	return {
		drawingIds,
		nextDrawingId: nextFreeNumericId(drawingIds, 1),
		bookmarkIds,
		nextBookmarkId: nextFreeNumericId(bookmarkIds, 0),
	};
}

function nextFreeNumericId(used: Set<string>, floor: number): number {
	let next = floor;
	for (const id of used) {
		const numeric = Number(id);
		if (Number.isFinite(numeric) && numeric >= next) next = numeric + 1;
	}
	return next;
}

/** The `r:` prefix (officeDocument/2006/relationships) exists solely to mark
 *  relationship-reference attributes, so ANY `r:*` attribute is an rId — a
 *  name shortlist would rot as DML/VML grow. A raw fragment naming an rId
 *  that doesn't exist is exactly the "unreadable content" corruption the
 *  relationships invariant guards against. */
function checkRelationshipReferences(
	document: Document,
	nodes: XmlNode[],
): void {
	for (const node of nodes) {
		walk(node, (element) => {
			for (const [key, value] of Object.entries(element.attributes)) {
				if (prefixOf(key) !== "r") continue;
				if (document.relationships.findByRid(value)) continue;
				throw new RawError(
					"INVALID_XML",
					`Fragment references relationship "${value}" (${key} on <${element.tag}>) which does not exist — a dangling rId corrupts the file`,
					"Create the target first (docx images add / docx hyperlinks add — or docx raw insert with a <Relationship Id Type Target/> root for anything else), then reference its rId — or drop the reference.",
				);
			}
		});
	}
}

function checkNoteAndCommentReferences(
	document: Document,
	nodes: XmlNode[],
): void {
	const checks: [string, (id: string) => boolean, string][] = [
		[
			"w:footnoteReference",
			(id) => document.footnotes?.findByNumericId(id) !== undefined,
			"footnote",
		],
		[
			"w:endnoteReference",
			(id) => document.endnotes?.findByNumericId(id) !== undefined,
			"endnote",
		],
		["w:commentReference", (id) => commentExists(document, id), "comment"],
		["w:commentRangeStart", (id) => commentExists(document, id), "comment"],
		["w:commentRangeEnd", (id) => commentExists(document, id), "comment"],
	];
	for (const node of nodes) {
		walk(node, (element) => {
			for (const [tag, exists, noun] of checks) {
				if (element.tag !== tag) continue;
				const id = element.getAttribute("w:id");
				if (id === undefined || exists(id)) continue;
				throw new RawError(
					"INVALID_XML",
					`Fragment references ${noun} id "${id}" (<${tag}>) which does not exist`,
					`Add the ${noun} first (docx ${noun}s add / docx comments add), or drop the reference.`,
				);
			}
		});
	}
}

function commentExists(document: Document, id: string): boolean {
	return document.comments?.hasId(id) ?? false;
}

/** `wp:docPr`/`pic:cNvPr` ids must be unique per document or Word flags
 *  corruption. A fragment copied from `raw get` (or another document) will
 *  collide — re-mint silently, continuing from the context's cursor. */
function remintDrawingIds(
	nodes: XmlNode[],
	context: ReferenceAuditContext,
): void {
	for (const node of nodes) {
		walk(node, (element) => {
			if (!isDrawingIdCarrier(element.tag)) return;
			const id = element.getAttribute("id");
			if (id === undefined || !context.drawingIds.has(id)) {
				if (id !== undefined) context.drawingIds.add(id);
				return;
			}
			const minted = String(context.nextDrawingId++);
			element.setAttribute("id", minted);
			context.drawingIds.add(minted);
		});
	}
}

/** Bookmark ids are document-unique; `w:bookmarkStart`/`w:bookmarkEnd` pair by
 *  id, so collisions re-map through one old→new table (scoped to the fragment,
 *  so its pairs stay linked). */
function remintBookmarkIds(
	nodes: XmlNode[],
	context: ReferenceAuditContext,
): void {
	const remapped = new Map<string, string>();
	for (const node of nodes) {
		walk(node, (element) => {
			if (!isBookmarkMarker(element.tag)) return;
			const id = element.getAttribute("w:id");
			if (id === undefined) return;
			const already = remapped.get(id);
			if (already !== undefined) {
				element.setAttribute("w:id", already);
				return;
			}
			if (!context.bookmarkIds.has(id)) {
				context.bookmarkIds.add(id);
				return;
			}
			const minted = String(context.nextBookmarkId++);
			remapped.set(id, minted);
			context.bookmarkIds.add(minted);
			element.setAttribute("w:id", minted);
		});
	}
}

function isBookmarkMarker(tag: string): boolean {
	return tag === "w:bookmarkStart" || tag === "w:bookmarkEnd";
}

const STYLE_REFERENCE_TAGS = new Set(["w:pStyle", "w:rStyle", "w:tblStyle"]);

function warnOnUnknownStyles(
	document: Document,
	nodes: XmlNode[],
	warnings: string[],
): void {
	const reported = new Set<string>();
	for (const node of nodes) {
		walk(node, (element) => {
			if (!STYLE_REFERENCE_TAGS.has(element.tag)) return;
			const styleId = element.getAttribute("w:val");
			if (styleId === undefined || reported.has(styleId)) return;
			if (document.styles?.hasStyle(styleId)) return;
			reported.add(styleId);
			warnings.push(
				`style "${styleId}" is not defined in styles.xml — Word will fall back to defaults (define it with: docx styles create)`,
			);
		});
	}
}

function warnOnUnknownNumbering(
	document: Document,
	nodes: XmlNode[],
	warnings: string[],
): void {
	const reported = new Set<string>();
	for (const node of nodes) {
		walk(node, (element) => {
			if (element.tag !== "w:numId") return;
			const numId = element.getAttribute("w:val");
			if (numId === undefined || numId === "0" || reported.has(numId)) return;
			if (document.numbering?.hasNum(numId)) return;
			reported.add(numId);
			warnings.push(
				`numbering id "${numId}" is not defined in numbering.xml — the list will not number`,
			);
		});
	}
}

function walk(node: XmlNode, visit: (element: XmlNode) => void): void {
	if (node.isText) return;
	visit(node);
	for (const child of node.children) walk(child, visit);
}
