import { w } from "../../jsx";
import { XmlNode } from "../../parser";
import type { ContentTypesView } from "./content-types";
import type { Pkg } from "./package";
import type { RelationshipsView } from "./relationships";

export const SETTINGS_PART_NAME = "word/settings.xml";
const SETTINGS_RELATIONSHIP_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
const SETTINGS_CONTENT_TYPE =
	"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";

const W_NAMESPACE =
	"http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** The document-level track-changes toggle: `<w:trackRevisions/>` is the real
 *  CT_Settings element (§17.15.1.90 — what Word writes); `w:trackChanges` is
 *  the misnamed element earlier docx-cli versions emitted (Word ignores it).
 *  Read both so those documents still report tracking-on; write only the real
 *  one and drop the legacy on toggle. */
const TRACK_TOGGLE_TAGS = new Set<string>([
	"w:trackRevisions",
	"w:trackChanges",
]);

/** The full CT_Settings child sequence (ECMA-376 §17.15.1.78), extracted from
 *  the bundled transitional `wml.xsd` — the settings analog of
 *  `PPR_CHILD_ORDER`/`SECTPR_CHILD_ORDER`. Every settings toggle splices via
 *  `insertSettingsChildInOrder`, so adding a new one is one call, not a new
 *  hand-maintained successor set. Extension elements Word appends after the
 *  schema sequence (`w14:docId`, `w15:docId`, …) are deliberately absent —
 *  unknown tags rank last, which is exactly where they live. */
const SETTINGS_CHILD_ORDER = [
	"w:writeProtection",
	"w:view",
	"w:zoom",
	"w:removePersonalInformation",
	"w:removeDateAndTime",
	"w:doNotDisplayPageBoundaries",
	"w:displayBackgroundShape",
	"w:printPostScriptOverText",
	"w:printFractionalCharacterWidth",
	"w:printFormsData",
	"w:embedTrueTypeFonts",
	"w:embedSystemFonts",
	"w:saveSubsetFonts",
	"w:saveFormsData",
	"w:mirrorMargins",
	"w:alignBordersAndEdges",
	"w:bordersDoNotSurroundHeader",
	"w:bordersDoNotSurroundFooter",
	"w:gutterAtTop",
	"w:hideSpellingErrors",
	"w:hideGrammaticalErrors",
	"w:activeWritingStyle",
	"w:proofState",
	"w:formsDesign",
	"w:attachedTemplate",
	"w:linkStyles",
	"w:stylePaneFormatFilter",
	"w:stylePaneSortMethod",
	"w:documentType",
	"w:mailMerge",
	"w:revisionView",
	"w:trackRevisions",
	"w:doNotTrackMoves",
	"w:doNotTrackFormatting",
	"w:documentProtection",
	"w:autoFormatOverride",
	"w:styleLockTheme",
	"w:styleLockQFSet",
	"w:defaultTabStop",
	"w:autoHyphenation",
	"w:consecutiveHyphenLimit",
	"w:hyphenationZone",
	"w:doNotHyphenateCaps",
	"w:showEnvelope",
	"w:summaryLength",
	"w:clickAndTypeStyle",
	"w:defaultTableStyle",
	"w:evenAndOddHeaders",
	"w:bookFoldRevPrinting",
	"w:bookFoldPrinting",
	"w:bookFoldPrintingSheets",
	"w:drawingGridHorizontalSpacing",
	"w:drawingGridVerticalSpacing",
	"w:displayHorizontalDrawingGridEvery",
	"w:displayVerticalDrawingGridEvery",
	"w:doNotUseMarginsForDrawingGridOrigin",
	"w:drawingGridHorizontalOrigin",
	"w:drawingGridVerticalOrigin",
	"w:doNotShadeFormData",
	"w:noPunctuationKerning",
	"w:characterSpacingControl",
	"w:printTwoOnOne",
	"w:strictFirstAndLastChars",
	"w:noLineBreaksAfter",
	"w:noLineBreaksBefore",
	"w:savePreviewPicture",
	"w:doNotValidateAgainstSchema",
	"w:saveInvalidXml",
	"w:ignoreMixedContent",
	"w:alwaysShowPlaceholderText",
	"w:doNotDemarcateInvalidXml",
	"w:saveXmlDataOnly",
	"w:useXSLTWhenSaving",
	"w:saveThroughXslt",
	"w:showXMLTags",
	"w:alwaysMergeEmptyNamespace",
	"w:updateFields",
	"w:hdrShapeDefaults",
	"w:footnotePr",
	"w:endnotePr",
	"w:compat",
	"w:docVars",
	"w:rsids",
	"m:mathPr",
	"w:attachedSchema",
	"w:themeFontLang",
	"w:clrSchemeMapping",
	"w:doNotIncludeSubdocsInStats",
	"w:doNotAutoCompressPictures",
	"w:forceUpgrade",
	"w:captions",
	"w:readModeInkLockDown",
	"w:smartTagType",
	"sl:schemaLibrary",
	"w:shapeDefaults",
	"w:doNotEmbedSmartTags",
	"w:decimalSymbol",
	"w:listSeparator",
] as const;

function settingsChildRank(tag: string): number {
	const index = SETTINGS_CHILD_ORDER.indexOf(
		tag as (typeof SETTINGS_CHILD_ORDER)[number],
	);
	return index >= 0 ? index : SETTINGS_CHILD_ORDER.length;
}

/** Splice a child into `<w:settings>` at its CT_Settings slot: before the
 *  first element that ranks after it. Unknown tags (extension elements like
 *  `w14:docId`) rank last, matching where Word writes them. */
function insertSettingsChildInOrder(root: XmlNode, child: XmlNode): void {
	const rank = settingsChildRank(child.tag);
	const at = root.children.findIndex(
		(existing) => !existing.isText && settingsChildRank(existing.tag) > rank,
	);
	if (at < 0) root.children.push(child);
	else root.children.splice(at, 0, child);
}

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
		return root.children.some((child) => TRACK_TOGGLE_TAGS.has(child.tag));
	}

	/** Toggle the document-level track-changes setting. The real CT_Settings
	 *  element is `<w:trackRevisions/>` (§17.15.1.90 — what Word itself writes;
	 *  `w:trackChanges` does not exist in the schema). Earlier docx-cli versions
	 *  emitted the misnamed `<w:trackChanges/>`, which Word ignores — so we READ
	 *  both (a doc we toggled still reads as tracking-on) and MIGRATE the legacy
	 *  element to the real one whenever the toggle runs. */
	setTrackChangesEnabled(on: boolean): void {
		const root = this.ensureSettingsRoot();
		if (!on) {
			root.children = root.children.filter(
				(child) => !TRACK_TOGGLE_TAGS.has(child.tag),
			);
			return;
		}
		root.children = root.children.filter(
			(child) => child.tag !== "w:trackChanges",
		);
		if (root.children.some((child) => child.tag === "w:trackRevisions")) {
			return;
		}
		insertSettingsChildInOrder(root, <w.trackRevisions />);
	}

	/** Ensure `<w:footnotePr>` / `<w:endnotePr>` is present, declaring the
	 *  reserved separator (id -1) + continuationSeparator (id 0) notes that live
	 *  in `footnotes.xml` / `endnotes.xml`. Word REQUIRES this settings-level
	 *  pointer to render a notes part — without it Word reports the document as
	 *  unreadable and "repairs" it by adding exactly this. Idempotent; spliced
	 *  at its CT_Settings slot so the child order stays valid regardless of
	 *  which neighbors the part already carries. */
	ensureNotePr(kind: "footnote" | "endnote"): void {
		const tag = kind === "footnote" ? "w:footnotePr" : "w:endnotePr";
		const root = this.ensureSettingsRoot();
		if (root.children.some((child) => child.tag === tag)) return;
		const NotePr = kind === "footnote" ? w.footnotePr : w.endnotePr;
		const Note = kind === "footnote" ? w.footnote : w.endnote;
		insertSettingsChildInOrder(
			root,
			<NotePr>
				<Note w-id="-1" />
				<Note w-id="0" />
			</NotePr>,
		);
	}

	/** Ensure `<w:evenAndOddHeaders/>` is present — the DOCUMENT-level toggle that
	 *  makes Word honor `even`-type header/footer references (without it, an even
	 *  marginal is ignored and the default applies to every page). Idempotent;
	 *  spliced at its CT_Settings slot. */
	ensureEvenAndOddHeaders(): void {
		const root = this.ensureSettingsRoot();
		if (root.children.some((child) => child.tag === "w:evenAndOddHeaders")) {
			return;
		}
		insertSettingsChildInOrder(root, <w.evenAndOddHeaders />);
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
