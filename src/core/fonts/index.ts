import type { Document } from "../ast/document";
import type { Pkg } from "../ast/document/package";
import type { RelationshipsView } from "../ast/document/relationships";
import { applyRunFont } from "../ast/document/styles";
import { XmlNode } from "../parser";

const THEME_RELATIONSHIP_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

/** Cross-cutting lens over "set the document's font." A document font lives in
 *  two parts at once — `word/styles.xml` `<w:docDefaults>` (the formal default)
 *  AND `word/theme/theme1.xml` `<a:fontScheme>` (what theme-referencing styles
 *  resolve through) — so setting only one silently loses to the other. `Fonts`
 *  writes both. The styles tree is owned by `StylesView` (saved on the normal
 *  path); the theme isn't a modeled view, so it's read, mutated, and staged back
 *  through `Pkg` here — only when this runs, never on unrelated saves.
 *
 *  Stateless; constructed at the call site: `await new Fonts(document).setDefault(name)`. */
export class Fonts {
	constructor(private document: Document) {}

	/** Make `fontName` the document default. Sets docDefaults rFonts + theme
	 *  major/minor `<a:latin>`, so body and theme-following headings both adopt it;
	 *  styles/runs that pin their OWN font are preserved (a code monospace, a
	 *  deliberately-Arial run) unless `opts.all`, which repoints every explicit
	 *  `<w:rFonts>` across styles, body, and notes for a guaranteed-uniform doc.
	 *  `opts.sizeHalfPoints` additionally sets the default size on the same
	 *  docDefaults write. The caller saves (so `--dry-run` can skip persistence). */
	async setDefault(
		fontName: string,
		opts: { sizeHalfPoints?: number; all?: boolean } = {},
	): Promise<SetDefaultFontResult> {
		const styles = this.document.ensureStyles();
		styles.setDefaultFont(fontName);
		if (opts.sizeHalfPoints !== undefined) {
			styles.setDefaultSizeHalfPoints(opts.sizeHalfPoints);
		}

		const themeUpdated = await applyThemeLatinFonts(
			this.document.pkg,
			themePartName(this.document.relationships),
			fontName,
		);

		// Computed BEFORE any override so it reflects styles that pin their own
		// font; the CLI reports them as "left untouched" (default) or "repointed"
		// (--all).
		const explicitStyles = styles.explicitFontStyleIds();

		let repointed = 0;
		if (opts.all) {
			repointed += styles.overrideStyleFonts(fontName);
			repointed += repointTreeFonts(this.document.documentTree, fontName);
			if (this.document.footnotes) {
				repointed += repointTreeFonts(this.document.footnotes.tree, fontName);
			}
			if (this.document.endnotes) {
				repointed += repointTreeFonts(this.document.endnotes.tree, fontName);
			}
		}

		return {
			font: fontName,
			sizeHalfPoints: opts.sizeHalfPoints,
			themeUpdated,
			explicitStyles,
			repointed,
			all: Boolean(opts.all),
		};
	}
}

export type SetDefaultFontResult = {
	font: string;
	sizeHalfPoints?: number;
	/** Whether a theme part existed and its font scheme was updated. */
	themeUpdated: boolean;
	/** Style ids that pin their own explicit font (override the default). */
	explicitStyles: string[];
	/** Count of explicit `<w:rFonts>` repointed across styles/body/notes (--all). */
	repointed: number;
	all: boolean;
};

/** The document's theme part path, resolved from the theme relationship's
 *  `Target` (relative to `word/`). Falls back to the conventional
 *  `word/theme/theme1.xml` when no theme relationship exists — Word, LibreOffice,
 *  and Pandoc all emit that path, but a hand-built package can point elsewhere
 *  (e.g. `theme/theme2.xml`), and the hardcoded path would silently no-op there. */
function themePartName(relationships: RelationshipsView): string {
	const rel = relationships
		.list()
		.find((info) => info.type === THEME_RELATIONSHIP_TYPE);
	if (!rel) return "word/theme/theme1.xml";
	return rel.target.startsWith("/")
		? rel.target.slice(1)
		: `word/${rel.target}`;
}

/** Repoint the theme's major + minor `<a:latin>` typefaces to `fontName` and
 *  stage the part back. Returns false (no-op) when there's no theme part or no
 *  font scheme. The theme isn't a modeled view, so we read/mutate/write it
 *  here through `Pkg` rather than on every `Document.save`. */
async function applyThemeLatinFonts(
	pkg: Pkg,
	partName: string,
	fontName: string,
): Promise<boolean> {
	const tree = await pkg.readPart(partName);
	if (!tree) return false;
	const fontScheme = XmlNode.findRoot(tree, "a:theme")
		?.findChild("a:themeElements")
		?.findChild("a:fontScheme");
	if (!fontScheme) return false;

	let changed = false;
	for (const slot of ["a:majorFont", "a:minorFont"] as const) {
		const latin = fontScheme.findChild(slot)?.findChild("a:latin");
		if (!latin) continue;
		latin.setAttribute("typeface", fontName);
		// A stale panose hint for the old font misleads Word's substitution.
		delete latin.attributes.panose;
		changed = true;
	}

	if (changed) pkg.writeText(partName, XmlNode.serialize(tree));
	return changed;
}

/** Repoint every `<w:rFonts>` anywhere in a part tree (body or a note part) to
 *  `fontName`, returning the count changed. The `--all` hammer — forces even
 *  runs that pinned their own font (e.g. inline code) onto the document font. */
function repointTreeFonts(tree: XmlNode[], fontName: string): number {
	let count = 0;
	const visit = (node: XmlNode): void => {
		if (node.tag === "w:rFonts") {
			applyRunFont(node, fontName);
			count++;
			return;
		}
		for (const child of node.children) visit(child);
	};
	for (const node of tree) visit(node);
	return count;
}
