import { XmlNode } from "../../parser";
import { buildBody } from "../read";
import type { Body, TrackedChangeReference } from "./body";
import { CommentsView } from "./comments";
import { ContentTypesView } from "./content-types";
import { CorePropertiesView } from "./core-properties";
import { NotesView } from "./notes";
import { NumberingView } from "./numbering";
import { Pkg } from "./package";
import { RelationshipsView } from "./relationships";
import { SettingsView } from "./settings";
import { StylesView } from "./styles";

export class Document {
	pkg: Pkg;
	documentTree: XmlNode[];
	comments?: CommentsView;
	footnotes?: NotesView;
	endnotes?: NotesView;
	numbering?: NumberingView;
	relationships: RelationshipsView;
	contentTypes: ContentTypesView;
	coreProperties?: CorePropertiesView;
	settings?: SettingsView;
	styles?: StylesView;

	body!: Body;
	trackedChangeReferences: Map<string, TrackedChangeReference> = new Map();

	constructor(init: {
		pkg: Pkg;
		documentTree: XmlNode[];
		comments?: CommentsView;
		footnotes?: NotesView;
		endnotes?: NotesView;
		numbering?: NumberingView;
		relationships: RelationshipsView;
		contentTypes: ContentTypesView;
		coreProperties?: CorePropertiesView;
		settings?: SettingsView;
		styles?: StylesView;
	}) {
		this.pkg = init.pkg;
		this.documentTree = init.documentTree;
		this.comments = init.comments;
		this.footnotes = init.footnotes;
		this.endnotes = init.endnotes;
		this.numbering = init.numbering;
		this.relationships = init.relationships;
		this.contentTypes = init.contentTypes;
		this.coreProperties = init.coreProperties;
		this.settings = init.settings;
		this.styles = init.styles;
	}

	static async open(path: string): Promise<Document> {
		const pkg = await Pkg.open(path);
		const view = new Document({
			pkg,
			documentTree: await pkg.ensurePart("word/document.xml"),
			relationships: await RelationshipsView.fromPackage(pkg),
			contentTypes: await ContentTypesView.fromPackage(pkg),
			coreProperties: await CorePropertiesView.fromPackage(pkg),
			settings: await SettingsView.fromPackage(pkg),
			styles: await StylesView.fromPackage(pkg),
			numbering: await NumberingView.fromPackage(pkg),
			footnotes: await NotesView.fromPackage(pkg, "footnote"),
			endnotes: await NotesView.fromPackage(pkg, "endnote"),
			comments: await CommentsView.fromPackage(pkg),
		});

		view.body = buildBody(view, pkg.path);
		return view;
	}

	/** Construct a synthetic Document from raw XML strings — no zip file, no
	 * disk read. Required parts (`relationshipsXml`, `contentTypesXml`)
	 * default to empty OOXML stubs. The Pkg is `Pkg.empty(path)`, so
	 * `view.save()` works (writes a zip containing whatever the views
	 * serialize) but `pkg.readBytes(...)` returns nothing — i.e., this is
	 * fine for reader / mutation tests, not for round-tripping media. */
	static fromXml(parts: {
		documentXml: string;
		relationshipsXml?: string;
		contentTypesXml?: string;
		settingsXml?: string;
		stylesXml?: string;
		numberingXml?: string;
		footnotesXml?: string;
		endnotesXml?: string;
		commentsXml?: string;
		commentsExtendedXml?: string;
		corePropertiesXml?: string;
		path?: string;
	}): Document {
		const path = parts.path ?? "synthetic.docx";
		const view = new Document({
			pkg: Pkg.empty(path),
			documentTree: XmlNode.parse(parts.documentXml),
			relationships: RelationshipsView.fromXml(parts.relationshipsXml),
			contentTypes: ContentTypesView.fromXml(parts.contentTypesXml),
			coreProperties: CorePropertiesView.fromXml(parts.corePropertiesXml),
			settings: SettingsView.fromXml(parts.settingsXml),
			styles: StylesView.fromXml(parts.stylesXml),
			numbering: NumberingView.fromXml(parts.numberingXml),
			footnotes: NotesView.fromXml("footnote", parts.footnotesXml),
			endnotes: NotesView.fromXml("endnote", parts.endnotesXml),
			comments: CommentsView.fromXml(
				parts.commentsXml,
				parts.commentsExtendedXml,
			),
		});
		view.body = buildBody(view, path);
		return view;
	}

	/** Rebuild the body AST + locator maps from the current (possibly mutated)
	 *  `documentTree`. Call after a structural mutation to read back the
	 *  freshly-assigned locators (e.g. the `pN` of a just-inserted block).
	 *  Block ids are positional, so they shift after inserts/deletes. */
	reread(): void {
		this.body = buildBody(this, this.body.path);
	}

	async save(path?: string): Promise<void> {
		this.pkg.writeText(
			"word/document.xml",
			XmlNode.serialize(this.documentTree),
		);
		this.relationships.writeTo(this.pkg);
		this.contentTypes.writeTo(this.pkg);
		this.coreProperties?.writeTo(this.pkg);
		this.settings?.writeTo(this.pkg);
		this.styles?.writeTo(this.pkg);
		this.numbering?.writeTo(this.pkg);
		this.footnotes?.writeTo(this.pkg);
		this.endnotes?.writeTo(this.pkg);
		this.comments?.writeTo(this.pkg);
		await this.pkg.save(path);
	}
	/** Convenience over `view.settings?.isTrackChangesEnabled() ?? false` —
	 * every mutating CLI verb checks this once per call to decide whether to
	 * handle change tracking. Returns false when the doc has no settings part
	 * at all (i.e. tracking can't be on). */
	isTrackChangesEnabled(): boolean {
		return this.settings?.isTrackChangesEnabled() ?? false;
	}

	/** Materialize the optional settings part if absent (mints rel +
	 * content-type via `SettingsView.register`); cached on the view. */
	ensureSettings(): SettingsView {
		if (this.settings) return this.settings;
		this.settings = SettingsView.register(this);
		return this.settings;
	}

	/** Materialize the optional styles part if absent (mints rel +
	 * content-type via `StylesView.register`); cached on the view. */
	ensureStyles(): StylesView {
		if (this.styles) return this.styles;
		this.styles = StylesView.register(this);
		return this.styles;
	}

	/** Materialize the optional numbering part if absent (mints rel +
	 * content-type via `NumberingView.register`); cached on the view. */
	ensureNumbering(): NumberingView {
		if (this.numbering) return this.numbering;
		this.numbering = NumberingView.register(this);
		return this.numbering;
	}

	/** Materialize the optional footnotes part if absent (mints rel +
	 * content-type via `NotesView.register`, seeded with the separator +
	 * continuationSeparator boilerplate Word expects); cached on the view. */
	ensureFootnotes(): NotesView {
		this.ensureNoteInfrastructure();
		// biome-ignore lint/style/noNonNullAssertion: set by ensureNoteInfrastructure.
		return this.footnotes!;
	}

	/** Materialize the optional endnotes part if absent — see `ensureFootnotes`. */
	ensureEndnotes(): NotesView {
		this.ensureNoteInfrastructure();
		// biome-ignore lint/style/noNonNullAssertion: set by ensureNoteInfrastructure.
		return this.endnotes!;
	}

	/** Provision the full footnote/endnote infrastructure Word requires whenever
	 * EITHER notes part is needed: both `footnotes.xml` AND `endnotes.xml` (Word
	 * always pairs them), plus the matching `<w:footnotePr>` / `<w:endnotePr>`
	 * separator declarations in `settings.xml`. Without those settings pointers
	 * Word reports "unreadable content" and repairs the file by adding exactly
	 * this. Idempotent — only registers what's missing. */
	private ensureNoteInfrastructure(): void {
		if (!this.footnotes) this.footnotes = NotesView.register("footnote", this);
		if (!this.endnotes) this.endnotes = NotesView.register("endnote", this);
		const settings = this.ensureSettings();
		settings.ensureNotePr("footnote");
		settings.ensureNotePr("endnote");
	}

	/** Materialize the optional comments part if absent (mints rel +
	 * content-type via `CommentsView.register`); cached on the view. */
	ensureComments(): CommentsView {
		if (this.comments) return this.comments;
		this.comments = CommentsView.register(this);
		return this.comments;
	}

	/** Materialize the optional commentsExtended part if absent — ensures the
	 * base comments part first, then mints the extended rel + content-type
	 * via `CommentsView.registerExtended`. */
	ensureCommentsExtended(): CommentsView {
		const comments = this.ensureComments();
		comments.registerExtended(this);
		return comments;
	}
}
