import { w } from "../../jsx";
import { XmlNode } from "../../parser";
import type { ContentTypesView } from "./content-types";
import type { Pkg } from "./package";
import type { RelationshipsView } from "./relationships";

const SETTINGS_PART_NAME = "word/settings.xml";
const SETTINGS_RELATIONSHIP_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
const SETTINGS_CONTENT_TYPE =
	"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";

const W_NAMESPACE =
	"http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** CT_Settings elements that follow `<w:footnotePr>`/`<w:endnotePr>` in the
 *  ECMA-376 §17.15.1.78 sequence — note-pr inserts before the first one present
 *  to stay schema-valid. Covers the tail elements that actually occur in real
 *  settings parts (Word always writes `<w:compat>`; the rest appear variably).
 *  Namespaced docId variants (`w14:`/`w15:`) are the common round-trip cases. */
const NOTE_PR_SUCCESSORS = new Set<string>([
	"w:compat",
	"w:rsids",
	"m:mathPr",
	"w:themeFontLang",
	"w:clrSchemeMapping",
	"w:shapeDefaults",
	"w:decimalSymbol",
	"w:listSeparator",
	"w:docId",
	"w14:docId",
	"w15:docId",
	"w:chartTrackingRefBased",
]);

/** CT_Settings elements that follow `<w:evenAndOddHeaders/>` (§17.15.1.78 pos 43)
 *  but PRECEDE `<w:footnotePr>`/`<w:endnotePr>` and the note-pr tail. The toggle
 *  must insert before the first of THESE too — otherwise on the canonical Word
 *  template (`…defaultTabStop, characterSpacingControl, compat…`) it would skip
 *  past `characterSpacingControl` (pos 57) and land out of order. Covers the
 *  mid-tail elements that realistically occur; Word/LibreOffice tolerate the rest. */
const EVEN_AND_ODD_SUCCESSORS = new Set<string>([
	"w:footnotePr",
	"w:endnotePr",
	"w:characterSpacingControl",
	"w:doNotUseMarginsForDrawingGridOrigin",
	"w:displayHangulFixedWidth",
	"w:noPunctuationKerning",
	"w:printTwoOnOne",
	"w:strictFirstAndLastChars",
	"w:savePreviewPicture",
	"w:updateFields",
	"w:hdrShapeDefaults",
	...NOTE_PR_SUCCESSORS,
]);

export class SettingsView {
	tree: XmlNode[];

	constructor(tree: XmlNode[]) {
		this.tree = tree;
	}

	/** Load this view from a package; returns undefined if the part is absent. */
	static async fromPackage(pkg: Pkg): Promise<SettingsView | undefined> {
		const tree = await pkg.readPart(SETTINGS_PART_NAME);
		return tree ? new SettingsView(tree) : undefined;
	}

	/** Parse a view from raw XML; returns undefined if the input is absent. */
	static fromXml(xml: string | undefined): SettingsView | undefined {
		return xml ? new SettingsView(XmlNode.parse(xml)) : undefined;
	}

	/** Serialize this view's tree into the package's `word/settings.xml`. */
	writeTo(pkg: Pkg): void {
		pkg.writeText(SETTINGS_PART_NAME, XmlNode.serialize(this.tree));
	}

	/** Mint the settings relationship + content-type override on the
	 * containing package and return a fresh empty view. Idempotent on the
	 * relationship target. Called by `Document.ensureSettings()`. */
	static register(deps: {
		relationships: RelationshipsView;
		contentTypes: ContentTypesView;
	}): SettingsView {
		if (!deps.relationships.hasTarget("settings.xml")) {
			deps.relationships.add(SETTINGS_RELATIONSHIP_TYPE, "settings.xml");
		}
		deps.contentTypes.registerPart(SETTINGS_PART_NAME, SETTINGS_CONTENT_TYPE);
		return new SettingsView([]);
	}

	isTrackChangesEnabled(): boolean {
		const root = XmlNode.findRoot(this.tree, "w:settings");
		if (!root) return false;
		return root.children.some((child) => child.tag === "w:trackChanges");
	}

	setTrackChangesEnabled(on: boolean): void {
		const root = this.ensureSettingsRoot();
		const hasTrackChanges = root.children.some(
			(child) => child.tag === "w:trackChanges",
		);
		if (on && !hasTrackChanges) {
			root.children.unshift(<w.trackChanges />);
		} else if (!on && hasTrackChanges) {
			root.children = root.children.filter(
				(child) => child.tag !== "w:trackChanges",
			);
		}
	}

	/** Ensure `<w:footnotePr>` / `<w:endnotePr>` is present, declaring the
	 *  reserved separator (id -1) + continuationSeparator (id 0) notes that live
	 *  in `footnotes.xml` / `endnotes.xml`. Word REQUIRES this settings-level
	 *  pointer to render a notes part — without it Word reports the document as
	 *  unreadable and "repairs" it by adding exactly this. Idempotent; inserted
	 *  before the first CT_Settings element that must follow note-pr so the child
	 *  order stays valid even when `<w:compat>` is absent (an imported settings
	 *  part may have none). Word/LibreOffice are lenient about settings order, so
	 *  the append fallback is a safe best-effort if no successor is present. */
	ensureNotePr(kind: "footnote" | "endnote"): void {
		const tag = kind === "footnote" ? "w:footnotePr" : "w:endnotePr";
		const root = this.ensureSettingsRoot();
		if (root.children.some((child) => child.tag === tag)) return;
		const NotePr = kind === "footnote" ? w.footnotePr : w.endnotePr;
		const Note = kind === "footnote" ? w.footnote : w.endnote;
		const node = (
			<NotePr>
				<Note w-id="-1" />
				<Note w-id="0" />
			</NotePr>
		);
		// `<w:footnotePr>` then `<w:endnotePr>` sit near the END of CT_Settings
		// (§17.15.1.78), just before this tail. Insert before the first tail
		// element present; when called for both kinds, footnotePr goes in first
		// and endnotePr then lands right after it (still before the tail), so
		// their required relative order is preserved.
		const successorIndex = root.children.findIndex((child) =>
			NOTE_PR_SUCCESSORS.has(child.tag),
		);
		if (successorIndex === -1) root.children.push(node);
		else root.children.splice(successorIndex, 0, node);
	}

	/** Ensure `<w:evenAndOddHeaders/>` is present — the DOCUMENT-level toggle that
	 *  makes Word honor `even`-type header/footer references (without it, an even
	 *  marginal is ignored and the default applies to every page). Idempotent.
	 *  In CT_Settings (§17.15.1.78) it sits before `<w:footnotePr>`/`<w:endnotePr>`
	 *  and the `<w:compat>` tail, so we insert before the first of those present
	 *  (Word/LibreOffice are lenient about settings order; append is a safe
	 *  fallback when none is present). */
	ensureEvenAndOddHeaders(): void {
		const root = this.ensureSettingsRoot();
		if (root.children.some((child) => child.tag === "w:evenAndOddHeaders")) {
			return;
		}
		const node = <w.evenAndOddHeaders />;
		const successorIndex = root.children.findIndex((child) =>
			EVEN_AND_ODD_SUCCESSORS.has(child.tag),
		);
		if (successorIndex === -1) root.children.push(node);
		else root.children.splice(successorIndex, 0, node);
	}

	/** Remove `<w:evenAndOddHeaders/>` — the counterpart to `ensureEvenAndOddHeaders`,
	 *  called when the last `even`-type header/footer is cleared (it's a document-
	 *  level toggle, so a stale one leaves even pages blank instead of inheriting
	 *  the default). No-op when absent. */
	removeEvenAndOddHeaders(): void {
		const root = XmlNode.findRoot(this.tree, "w:settings");
		if (!root) return;
		root.children = root.children.filter(
			(child) => child.tag !== "w:evenAndOddHeaders",
		);
	}

	private ensureSettingsRoot(): XmlNode {
		const existing = XmlNode.findRoot(this.tree, "w:settings");
		if (existing) return existing;
		const fresh = <w.settings {...{ "xmlns:w": W_NAMESPACE }} />;
		this.tree.push(fresh);
		return fresh;
	}
}
