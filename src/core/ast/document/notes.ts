import { type NoteKind, noteConfig } from "../../notes/config";
import { buildEmptyNotesTree } from "../../notes/empty";
import { XmlNode } from "../../parser";
import type { Note } from "../types";
import type { ContentTypesView } from "./content-types";
import type { Pkg } from "./package";
import type { RelationshipsView } from "./relationships";
import type { StylesView } from "./styles";

export class NotesView {
	kind: NoteKind;
	tree: XmlNode[];

	constructor(kind: NoteKind, tree: XmlNode[]) {
		this.kind = kind;
		this.tree = tree;
	}

	/** Load this view from a package; returns undefined if the part is absent. */
	static async fromPackage(
		pkg: Pkg,
		kind: NoteKind,
	): Promise<NotesView | undefined> {
		const tree = await pkg.readPart(noteConfig(kind).partName);
		return tree ? new NotesView(kind, tree) : undefined;
	}

	/** Parse a view from raw XML; returns undefined if the input is absent. */
	static fromXml(
		kind: NoteKind,
		xml: string | undefined,
	): NotesView | undefined {
		return xml ? new NotesView(kind, XmlNode.parse(xml)) : undefined;
	}

	/** Serialize this view's tree into the package's `word/${kind}s.xml`. */
	writeTo(pkg: Pkg): void {
		pkg.writeText(noteConfig(this.kind).partName, XmlNode.serialize(this.tree));
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
	 * body) and collapses note-body whitespace to single spaces. */
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
			const text = child.collectText().replace(/\s+/g, " ").trim();
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
