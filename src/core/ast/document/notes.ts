import { declareUsedNamespaces } from "../../namespaces";
import { type NoteKind, noteConfig } from "../../notes/config";
import { buildEmptyNotesTree } from "../../notes/empty";
import { XmlNode } from "../../parser";
import type { Note } from "../types";
import type { ContentTypesView } from "./content-types";
import type { Pkg } from "./package";
import { RelationshipsView, relsPartNameFor } from "./relationships";
import type { StylesView } from "./styles";

export class NotesView {
	kind: NoteKind;
	tree: XmlNode[];
	/** This part's own `word/_rels/<kind>s.xml.rels`, holding the hyperlink (and
	 *  any future media) relationships referenced from note bodies. Lazily
	 *  created the first time a body needs one; absent for plain-text notes. */
	relationships?: RelationshipsView;

	constructor(kind: NoteKind, tree: XmlNode[]) {
		this.kind = kind;
		this.tree = tree;
	}

	/** Load this view from a package; returns undefined if the part is absent.
	 *  Also attaches the part's own rels (`word/_rels/<kind>s.xml.rels`) when
	 *  present so note-body hyperlinks survive a round-trip. */
	static async fromPackage(
		pkg: Pkg,
		kind: NoteKind,
	): Promise<NotesView | undefined> {
		const tree = await pkg.readPart(noteConfig(kind).partName);
		if (!tree) return undefined;
		const view = new NotesView(kind, tree);
		const relsTree = await pkg.readPart(
			relsPartNameFor(noteConfig(kind).partName),
		);
		if (relsTree) {
			view.relationships = new RelationshipsView(
				relsTree,
				relsPartNameFor(noteConfig(kind).partName),
			);
		}
		return view;
	}

	/** This part's rels view, created empty (targeting the right `.rels` part)
	 *  on first use. Note-body hyperlink rels go here, not the document's. */
	ensureRelationships(): RelationshipsView {
		if (!this.relationships) {
			this.relationships = RelationshipsView.fromXml(
				undefined,
				relsPartNameFor(noteConfig(this.kind).partName),
			);
		}
		return this.relationships;
	}

	/** Parse a view from raw XML; returns undefined if the input is absent. */
	static fromXml(
		kind: NoteKind,
		xml: string | undefined,
	): NotesView | undefined {
		return xml ? new NotesView(kind, XmlNode.parse(xml)) : undefined;
	}

	/** Serialize this view's tree into the package's `word/${kind}s.xml`, plus
	 *  its rels part when any note-body relationships were minted. */
	writeTo(pkg: Pkg): void {
		declareUsedNamespaces(this.tree);
		pkg.writeText(noteConfig(this.kind).partName, XmlNode.serialize(this.tree));
		if (this.relationships?.hasAny()) this.relationships.writeTo(pkg);
	}

	/** Mint the relationship + content-type override for `word/${kind}s.xml`
	 * on the containing package and return a fresh view seeded with the
	 * separator + continuationSeparator boilerplate Word expects. Idempotent
	 * on the relationship target. Called by `Document.ensureFootnotes()` /
	 * `Document.ensureEndnotes()`. */
	static register(
		kind: NoteKind,
		deps: {
			relationships: RelationshipsView;
			contentTypes: ContentTypesView;
		},
	): NotesView {
		const config = noteConfig(kind);
		if (!deps.relationships.hasTarget(config.target)) {
			deps.relationships.add(config.relationshipType, config.target);
		}
		deps.contentTypes.registerPart(config.partName, config.contentType);
		return new NotesView(kind, buildEmptyNotesTree(kind));
	}

	/** Parse this part's tree into the AST `Footnote[]` the reader exposes on
	 * `Body.footnotes` / `Body.endnotes`. Skips Word's reserved separator /
	 * continuationSeparator entries (`w:type` set, never referenced from the
	 * body) and renders line breaks / tabs as `\n` / `\t` (via
	 * `collectTextWithBreaks`) so a `\n`-authored body round-trips. */
	toNotes(): Note[] {
		const root = this.findRoot();
		if (!root) return [];
		const config = noteConfig(this.kind);
		const out: Note[] = [];
		for (const child of root.children) {
			if (child.tag !== config.itemTag) continue;
			if (child.getAttribute("w:type")) continue;
			const numericId = child.getAttribute("w:id");
			if (numericId == null) continue;
			// Break/tab-aware so a body authored with `\n`/`\t` reads back faithfully
			// (the markdown footnote renderer collapses `\s+` for its single-line
			// definition, so the AST stays the lossless view without breaking GFM).
			const text = child.collectTextWithBreaks().trim();
			out.push({ id: `${config.idPrefix}${numericId}`, text });
		}
		return out;
	}

	listIds(): string[] {
		const root = this.findRoot();
		if (!root) return [];
		const config = noteConfig(this.kind);
		const out: string[] = [];
		for (const child of root.children) {
			if (child.tag !== config.itemTag) continue;
			const id = child.getAttribute("w:id");
			if (id) out.push(id);
		}
		return out;
	}

	/** Word reserves -1 (separator) and 0 (continuationSeparator) as the
	 *  boilerplate ids; user notes start at 1. We allocate `max(existing user
	 *  id) + 1` so we never collide with the reserved entries even if the
	 *  document was authored by something that used different defaults. */
	nextId(): string {
		const root = this.findRoot();
		if (!root) return "1";
		const config = noteConfig(this.kind);
		let highest = 0;
		for (const child of root.children) {
			if (child.tag !== config.itemTag) continue;
			const idAttribute = child.getAttribute("w:id");
			if (idAttribute == null) continue;
			const numeric = Number(idAttribute);
			if (Number.isFinite(numeric) && numeric > highest) highest = numeric;
		}
		return String(highest + 1);
	}

	findByNumericId(
		numericId: string,
	): { node: XmlNode; parent: XmlNode[] } | undefined {
		const root = this.findRoot();
		if (!root) return undefined;
		const config = noteConfig(this.kind);
		for (const child of root.children) {
			if (child.tag !== config.itemTag) continue;
			if (child.getAttribute("w:id") !== numericId) continue;
			return { node: child, parent: root.children };
		}
		return undefined;
	}

	/** A real note (not Word's `w:type` separator boilerplate) by numeric id. */
	findUserNote(numericId: string): XmlNode | undefined {
		const found = this.findByNumericId(numericId);
		if (!found || found.node.getAttribute("w:type")) return undefined;
		return found.node;
	}

	/** Copy note `numericId` under a fresh id (placed right after the original)
	 *  and return the new id — what Word does when text holding a note reference
	 *  is pasted. Markers whose ids must stay document-unique (bookmarks, comment
	 *  ranges + references, `w14` paragraph ids) are dropped from the copy, so the
	 *  clone can't collide with the original. The caller has already checked the
	 *  note exists (`findUserNote`). */
	cloneNote(numericId: string): string {
		const found = this.findByNumericId(numericId);
		if (!found) throw new Error(`Internal: no note ${numericId} to clone`);
		const copy = found.node.clone();
		const id = this.nextId();
		copy.setAttribute("w:id", id);
		stripUniqueMarkers(copy);
		found.parent.splice(found.parent.indexOf(found.node) + 1, 0, copy);
		return id;
	}

	/** Lazily provision the two character/paragraph styles a note relies on.
	 *  The baseline catalog defines both — these calls just register the style
	 *  nodes in `styles.xml` if not already present, which avoids Word's
	 *  fall-back to Normal (no superscript on the marker, default font on the
	 *  body). Takes the StylesView as a dependency (cross-view dep is explicit
	 *  rather than a back-ref to Document). */
	ensureNoteStyles(styles: StylesView): void {
		const config = noteConfig(this.kind);
		styles.ensureStyle(config.referenceStyle);
		styles.ensureStyle(config.textStyle);
	}

	private findRoot(): XmlNode | undefined {
		const config = noteConfig(this.kind);
		return XmlNode.findRoot(this.tree, config.rootTag);
	}
}

/** Drop, in place, every marker under `node` whose id must be unique across the
 *  document — so a cloned note body can't duplicate one. */
function stripUniqueMarkers(node: XmlNode): void {
	delete node.attributes["w14:paraId"];
	delete node.attributes["w14:textId"];
	node.children = node.children.filter((child) => !isUniqueMarker(child));
	for (const child of node.children) stripUniqueMarkers(child);
}

function isUniqueMarker(node: XmlNode): boolean {
	if (UNIQUE_MARKER_TAGS.has(node.tag)) return true;
	return (
		node.tag === "w:r" && node.findChild("w:commentReference") !== undefined
	);
}

const UNIQUE_MARKER_TAGS: ReadonlySet<string> = new Set([
	"w:bookmarkStart",
	"w:bookmarkEnd",
	"w:commentRangeStart",
	"w:commentRangeEnd",
]);
