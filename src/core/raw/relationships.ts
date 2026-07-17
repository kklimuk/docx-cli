import type { Document } from "../ast/document";
import { XmlNode } from "../parser";
import { RawError } from "./error";
import { assertWellFormed, parseRoots } from "./parse";

/** Raw addressing for the document's relationships part
 *  (`word/_rels/document.xml.rels`). Relationships are the other half of many
 *  escape-hatch constructs — an OLE object, chart, or external link in a body
 *  fragment references an rId that must exist — so `raw get/insert/replace`
 *  address them alongside blocks: `--at rIdN` targets one `<Relationship>`,
 *  `--at rels` (get only) prints the whole part, and an insert whose fragment
 *  roots are `<Relationship>` routes here by shape (no placement flags — the
 *  part is an unordered set keyed by Id).
 *
 *  The gate pipeline is proportionally smaller because the rels schema is
 *  tiny and NOT MCE-capable (no `mc:Ignorable`, so no `dcx:raw` marker, and
 *  rels never surface in `read` — they're plumbing, not content):
 *   1. well-formedness — same gate 1 as blocks
 *   2. `<Relationship>` roots only, no children, no unknown attributes
 *   3. Id semantics — insert: free id required (omit to have one minted);
 *      replace: the Id must keep the target's (a rename would dangle every
 *      body reference)
 *   4. Target integrity — an internal Target must resolve to an existing
 *      part (a relationship to a missing part is the same "unreadable
 *      content" class as a dangling rId, just pointed the other way) */
export type RelationshipMode =
	| { kind: "insert" }
	| { kind: "replace"; rId: string };

export type PreparedRelationships = {
	mode: RelationshipMode;
	nodes: XmlNode[];
	/** The Id of each prepared node, in order — minted ones included. */
	ids: string[];
};

/** Shape-sniff for CLI routing: does this fragment start with a
 *  `<Relationship>` root? Routing only — the gates still enforce. */
export function isRelationshipFragment(xml: string): boolean {
	return /^\s*<Relationship[\s/>]/.test(xml);
}

export function prepareRelationshipsFragment(
	document: Document,
	xml: string,
	mode: RelationshipMode,
): PreparedRelationships {
	assertWellFormed(xml);
	const nodes = parseRoots(xml, "relationships");
	if (mode.kind === "replace" && nodes.length !== 1) {
		throw new RawError(
			"INVALID_XML",
			"Replacing a relationship takes exactly one <Relationship> root",
		);
	}
	for (const node of nodes) {
		validateRelationshipShape(node);
		checkInternalTarget(document, node);
	}
	const ids = assignRelationshipIds(document, nodes, mode);
	return { mode, nodes, ids };
}

/** Splice a prepared fragment into the rels part. Replace removes the old
 *  node first (same Id, so body references stay valid); both paths re-index
 *  so the image/hyperlink lookup maps stay coherent with the tree. */
export function applyRelationshipsFragment(
	document: Document,
	prepared: PreparedRelationships,
): void {
	const root = XmlNode.findRoot(document.relationships.tree, "Relationships");
	if (!root) throw new Error("missing <Relationships> root");
	if (prepared.mode.kind === "replace") {
		document.relationships.remove(prepared.mode.rId);
	}
	root.children.push(...prepared.nodes);
	document.relationships.index(document.contentTypes);
}

/** The read half: one relationship's exact XML, or the whole part for
 *  `--at rels` (the discovery path — ids, types, targets in one look).
 *  Returns undefined when `rId` names nothing. */
export function serializeRelationships(
	document: Document,
	rId?: string,
): string | undefined {
	if (rId === undefined) {
		const root = XmlNode.findRoot(document.relationships.tree, "Relationships");
		return root ? XmlNode.serialize([root]) : undefined;
	}
	const node = document.relationships.findByRid(rId);
	return node ? XmlNode.serialize([node]) : undefined;
}

const RELS_NAMESPACE =
	"http://schemas.openxmlformats.org/package/2006/relationships";

const RELATIONSHIP_ATTRIBUTE_NAMES = new Set([
	"Id",
	"Type",
	"Target",
	"TargetMode",
]);

function validateRelationshipShape(node: XmlNode): void {
	const children = node.children.filter(
		(child) => !(child.isText && (child.text ?? "").trim() === ""),
	);
	if (children.length > 0) {
		throw new RawError(
			"INVALID_XML",
			"<Relationship> is an empty element — it takes attributes only",
		);
	}
	for (const key of Object.keys(node.attributes)) {
		if (RELATIONSHIP_ATTRIBUTE_NAMES.has(key)) continue;
		// A redundant default-namespace declaration matching the rels part is
		// harmless intent — strip it (the part root already declares it).
		if (key === "xmlns" && node.getAttribute("xmlns") === RELS_NAMESPACE) {
			node.attributes = Object.fromEntries(
				Object.entries(node.attributes).filter(([name]) => name !== "xmlns"),
			);
			continue;
		}
		throw new RawError(
			"INVALID_XML",
			`<Relationship> does not take a "${key}" attribute — only Id, Type, Target, TargetMode`,
		);
	}
	if (!node.getAttribute("Type")) {
		throw new RawError(
			"INVALID_XML",
			"<Relationship> requires a Type attribute (a relationship-type URI)",
		);
	}
	if (!node.getAttribute("Target")) {
		throw new RawError(
			"INVALID_XML",
			"<Relationship> requires a Target attribute",
		);
	}
	const targetMode = node.getAttribute("TargetMode");
	if (
		targetMode !== undefined &&
		targetMode !== "External" &&
		targetMode !== "Internal"
	) {
		throw new RawError(
			"INVALID_XML",
			`TargetMode must be "External" or "Internal" (the default) — got "${targetMode}"`,
		);
	}
}

/** An internal Target must resolve to an existing part. This is the dangling
 *  rId invariant pointed the other way: the body gate rejects a reference to
 *  a missing relationship; this gate rejects a relationship to a missing
 *  part. External targets (URLs) are unchecked by design. */
function checkInternalTarget(document: Document, node: XmlNode): void {
	if (node.getAttribute("TargetMode") === "External") return;
	const target = node.getAttribute("Target");
	if (!target) return;
	const partName = resolveRelativePart(target);
	if (document.pkg.hasPart(partName)) return;
	throw new RawError(
		"INVALID_XML",
		`Target "${target}" resolves to part "${partName}", which does not exist in the package — a relationship to a missing part corrupts the file the same way a dangling rId does`,
		`For a URL, add TargetMode="External". Create the part first: docx raw part add FILE --name ${partName} (--from bytes | --xml '…') — or docx images add for pictures.`,
	);
}

/** Resolve a rels Target against its base (`word/`, since this is the
 *  document part's rels): `media/image1.png` → `word/media/image1.png`,
 *  `../customXml/item1.xml` → `customXml/item1.xml`, a leading `/` is
 *  package-absolute. */
function resolveRelativePart(target: string): string {
	const base = target.startsWith("/") ? target.slice(1) : `word/${target}`;
	const segments: string[] = [];
	for (const segment of base.split("/")) {
		if (segment === "" || segment === ".") continue;
		if (segment === "..") {
			segments.pop();
			continue;
		}
		segments.push(segment);
	}
	return segments.join("/");
}

/** Insert: every explicit Id must be free (in the part AND the fragment);
 *  a node without an Id gets the next `rIdN` minted. Replace: the Id must
 *  stay the target's — stamped when omitted, rejected when different (a
 *  changed Id is a rename, which would dangle every body reference). */
function assignRelationshipIds(
	document: Document,
	nodes: XmlNode[],
	mode: RelationshipMode,
): string[] {
	if (mode.kind === "replace") {
		const node = nodes[0];
		if (!node) throw new Error("unreachable: replace takes one node");
		const id = node.getAttribute("Id");
		if (id !== undefined && id !== mode.rId) {
			throw new RawError(
				"INVALID_XML",
				`The replacement's Id="${id}" must keep the target's id ${mode.rId} — changing it would dangle every reference to ${mode.rId}`,
				"To rename, insert a new <Relationship> and update the body references to point at it.",
			);
		}
		if (id === undefined) node.setAttribute("Id", mode.rId);
		return [mode.rId];
	}
	const taken = new Set<string>();
	for (const node of nodes) {
		const id = node.getAttribute("Id");
		if (id === undefined) continue;
		if (document.relationships.findByRid(id)) {
			throw new RawError(
				"INVALID_XML",
				`Relationship "${id}" already exists — ids are unique within the part`,
				"Omit Id to have a free one minted, or pick another (docx raw get FILE --at rels lists them). To modify the existing one, use raw replace --at " +
					id +
					".",
			);
		}
		if (taken.has(id)) {
			throw new RawError(
				"INVALID_XML",
				`Duplicate Id="${id}" within the fragment`,
			);
		}
		taken.add(id);
	}
	let cursor = Number(document.relationships.nextId().slice("rId".length));
	const ids: string[] = [];
	for (const node of nodes) {
		const explicit = node.getAttribute("Id");
		if (explicit !== undefined) {
			ids.push(explicit);
			continue;
		}
		while (taken.has(`rId${cursor}`)) cursor++;
		const minted = `rId${cursor++}`;
		node.setAttribute("Id", minted);
		taken.add(minted);
		ids.push(minted);
	}
	return ids;
}
