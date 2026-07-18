import type { Document } from "../ast/document";
import type { ContentTypesView } from "../ast/document/content-types";
import { NUMBERING_PART_NAME, NumberingView } from "../ast/document/numbering";
import type { Pkg } from "../ast/document/package";
import { SETTINGS_PART_NAME, SettingsView } from "../ast/document/settings";
import { STYLES_PART_NAME, StylesView } from "../ast/document/styles";
import { isMarginalPartName } from "../marginals/config";
import { XmlNode } from "../parser";
import { RawError } from "./error";
import { assertWellFormed } from "./parse";

/** Raw addressing for OPC parts — the third address family after body blocks
 *  and relationships, surfaced as `docx raw part list|get|add|replace|edit
 *  --name NAME` (parts are the package's FILES, not document content, so they
 *  deliberately do NOT ride the `--at` locator system). `list`/`get` read any
 *  part (reading is always safe), `add` creates a NEW part (the missing half
 *  of the embedded-objects workflow: part → relationship → body fragment),
 *  and replace/edit swap or patch a whole XML part — view-owned parts route
 *  through their views via `partReplaceRoute` (a raw byte swap of those would
 *  be clobbered by the view's serialization on save).
 *
 *  Part names are the stable handle (they never shift), which is what keeps
 *  the write-read loop honest without inventing a sub-part locator grammar:
 *  inside a part there is deliberately NO positional addressing — a keyed
 *  construct (styleId, numId, …) belongs to its modeled verb, and anything
 *  else is whole-part get → modify → replace. */
export type PreparedPartAdd = {
	name: string;
	content: string | Uint8Array;
	/** XML parts only — the parsed root, for the CLI's schema gate. */
	root?: XmlNode;
	/** An Override to register; undefined when an extension Default covers it. */
	contentType?: string;
};

export type PreparedPartReplace = {
	name: string;
	xml: string;
	oldRoot: XmlNode;
	newRoot: XmlNode;
	route: PartReplaceRoute;
};

/** How a part replacement lands. `pkg`: unmodeled — stage the bytes, no view
 *  will fight them. `view`: the part is view-owned but self-contained enough
 *  to swap safely — the patched XML is reparsed INTO the owning view, so the
 *  view's serialization on save writes the patch instead of clobbering it.
 *  Parts whose ids pair with body markers (notes, comments) or whose content
 *  IS the body stay rejected toward their own surfaces. */
type PartReplaceRoute =
	| { kind: "pkg" }
	| { kind: "view"; apply: (document: Document, xml: string) => void };

/** Every OPC part with its content type — the `part list` discovery listing.
 *  Takes only the package transport + content-types view (never a `Document`):
 *  listing parts must not pay for the body-AST build. */
export function listPackageParts(
	pkg: Pkg,
	contentTypes: ContentTypesView,
): { name: string; contentType: string }[] {
	return pkg
		.listParts()
		.filter((name) => !name.endsWith("/"))
		.sort()
		.map((name) => ({
			name,
			contentType: contentTypes.lookupContentType(name),
		}));
}

export function preparePartAdd(
	document: Document,
	rawName: string,
	source: { xml?: string; bytes?: Uint8Array },
	explicitContentType: string | undefined,
): PreparedPartAdd {
	const name = normalizePartName(rawName);
	if (name.endsWith(".rels")) {
		throw new RawError(
			"USAGE",
			"Relationship parts aren't created directly — relationships are addressed as rels/rIdN",
			"Insert a <Relationship …/> fragment with raw insert; the rels part exists already.",
		);
	}
	if (document.pkg.hasPart(name)) {
		throw new RawError(
			"USAGE",
			`Part "${name}" already exists`,
			isXmlPartName(name)
				? `Replace it instead: docx raw part replace FILE --name ${name} --xml '…'.`
				: "Binary parts can't be replaced raw yet — docx images replace covers pictures.",
		);
	}
	if (source.xml !== undefined && !isXmlPartName(name)) {
		throw new RawError(
			"USAGE",
			`"${name}" is not an XML part name — pass the bytes with --from PATH`,
		);
	}
	if (source.xml === undefined && isXmlPartName(name)) {
		throw new RawError(
			"USAGE",
			`"${name}" is an XML part — pass its markup with --xml or --xml-file so it can be gated`,
		);
	}
	const contentType = resolveContentType(document, name, explicitContentType);
	if (source.xml !== undefined) {
		const root = parsePartRoot(source.xml);
		return { name, content: source.xml, root, contentType };
	}
	if (source.bytes === undefined) {
		throw new RawError("USAGE", "Missing part content");
	}
	return { name, content: source.bytes, contentType };
}

export function applyPartAdd(
	document: Document,
	prepared: PreparedPartAdd,
): void {
	if (typeof prepared.content === "string") {
		document.pkg.writeText(prepared.name, prepared.content);
	} else {
		document.pkg.writeBytes(prepared.name, prepared.content);
	}
	if (prepared.contentType) {
		document.contentTypes.registerPart(prepared.name, prepared.contentType);
	}
}

/** Whole-part replace for XML parts. On the pkg route the replacement is
 *  stored VERBATIM (leading XML declaration and all); on the view route it
 *  reparses into the owning view and lands through the view's serializer.
 *  Either way the gates parse a stripped copy, and the root tag must match
 *  the existing part's — a part's root IS its identity to every consumer
 *  that resolves it. `currentXml` lets `part edit` (which already read the
 *  part for its find/replace) pass it in and skip the second inflate. */
export async function preparePartReplace(
	document: Document,
	rawName: string,
	xml: string,
	currentXml?: string,
): Promise<PreparedPartReplace> {
	const name = normalizePartName(rawName);
	const route = partReplaceRoute(name);
	if ("hint" in route) {
		throw new RawError(
			"USAGE",
			`Part "${name}" can't be replaced raw — its content pairs with the body (or IS the body), so a whole-part swap could dangle references no gate can re-mint`,
			route.hint,
		);
	}
	if (!isXmlPartName(name)) {
		throw new RawError(
			"USAGE",
			`"${name}" is a binary part — only XML parts can be replaced raw`,
			"For pictures, docx images replace swaps the bytes safely.",
		);
	}
	const oldRoot = parsePartRoot(
		currentXml ?? (await document.pkg.readText(name)),
	);
	const newRoot = parsePartRoot(xml);
	if (oldRoot.tag !== newRoot.tag) {
		throw new RawError(
			"INVALID_XML",
			`The replacement's root <${newRoot.tag}> must keep the part's root <${oldRoot.tag}> — a part's root element is its identity`,
		);
	}
	return { name, xml, oldRoot, newRoot, route };
}

export function applyPartReplace(
	document: Document,
	prepared: PreparedPartReplace,
): void {
	if (prepared.route.kind === "view") {
		prepared.route.apply(document, prepared.xml);
		return;
	}
	document.pkg.writeText(prepared.name, prepared.xml);
}

/** OPC part names: no leading slash (we normalize it away), no empty/dot
 *  segments. Case is preserved. */
export function normalizePartName(rawName: string): string {
	const name = rawName.startsWith("/") ? rawName.slice(1) : rawName;
	if (name === "") throw new RawError("USAGE", "Empty part name");
	for (const segment of name.split("/")) {
		if (segment === "" || segment === "." || segment === "..") {
			throw new RawError(
				"USAGE",
				`Invalid part name "${rawName}" — use a plain package path like word/embeddings/object1.bin`,
			);
		}
	}
	return name;
}

export function isXmlPartName(name: string): boolean {
	return name.endsWith(".xml") || name.endsWith(".rels");
}

/** Route a part name to its replacement mechanism. View-owned parts whose
 *  content is SELF-CONTAINED (styles, numbering, settings, header/footer
 *  bodies) take the view route: the patched XML reparses into a fresh view
 *  (each id family they define is one Word tolerates dangling — a missing
 *  style/numId falls back to defaults, per the fragment gate's own warning
 *  policy). Parts whose ids PAIR with body markers (notes, comments — a
 *  vanished id there is the repair-prompt class), the body itself, rels, and
 *  the package plumbing reject toward their real surfaces.
 *
 *  Each view's own `PART_NAME` constant and `isMarginalPartName` are the
 *  authorities on which path a name is (not strings re-typed here), so a view
 *  that ever renames or adds a part can't drift out of this table silently. */
function partReplaceRoute(name: string): PartReplaceRoute | { hint: string } {
	const rejected: Record<string, string> = {
		"word/document.xml":
			"Address the body through block locators: raw get/insert/replace/edit --at pN|tN|sN.",
		"word/footnotes.xml":
			"Note ids pair with body reference marks — use docx footnotes add/edit/delete.",
		"word/endnotes.xml":
			"Note ids pair with body reference marks — use docx endnotes add/edit/delete.",
		"word/comments.xml":
			"Comment ids pair with body range marks — use docx comments add/reply/resolve.",
		"word/commentsExtended.xml":
			"Threading pairs with comments.xml — use docx comments add/reply/resolve.",
		"docProps/core.xml": "Core properties are managed by docx-cli.",
		"[Content_Types].xml":
			"Content types are registered automatically when parts are added.",
	};
	const hint = rejected[name];
	if (hint) return { hint };
	if (name.endsWith(".rels")) {
		return {
			hint: "Relationships are addressed directly: raw get --at rels, insert a <Relationship …/>, replace --at rIdN.",
		};
	}
	if (name === STYLES_PART_NAME) {
		return {
			kind: "view",
			apply: (document, xml) => {
				document.styles = StylesView.fromXml(xml);
			},
		};
	}
	if (name === NUMBERING_PART_NAME) {
		return {
			kind: "view",
			apply: (document, xml) => {
				document.numbering = NumberingView.fromXml(xml);
			},
		};
	}
	if (name === SETTINGS_PART_NAME) {
		return {
			kind: "view",
			apply: (document, xml) => {
				document.settings = SettingsView.fromXml(xml);
			},
		};
	}
	if (isMarginalPartName(name)) {
		return {
			kind: "view",
			apply: (document, xml) => {
				if (!document.marginals) {
					throw new Error(`no marginals view holds ${name}`);
				}
				document.marginals.setPart(name, XmlNode.parse(xml));
			},
		};
	}
	return { kind: "pkg" };
}

/** Explicit `--content-type` registers an Override; otherwise the part's
 *  extension must already have a Default in [Content_Types].xml — an OPC part
 *  without a content type makes Word reject the whole package, so this is a
 *  hard gate, not a guess. */
function resolveContentType(
	document: Document,
	name: string,
	explicit: string | undefined,
): string | undefined {
	if (explicit) return explicit;
	const extension = name.split(".").pop()?.toLowerCase() ?? "";
	if (extension && document.contentTypes.hasDefault(extension)) {
		return undefined;
	}
	throw new RawError(
		"USAGE",
		`No content type known for "${name}" — pass --content-type`,
		'Common ones: OLE object "application/vnd.openxmlformats-officedocument.oleObject"; embedded .docx "application/vnd.openxmlformats-officedocument.wordprocessingml.document"; chart XML "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"; plain XML "application/xml". Pictures are better served by docx images add.',
	);
}

/** Parse a whole part's XML: tolerate (and preserve — storage is verbatim) a
 *  leading XML declaration, require well-formedness and exactly one root. */
function parsePartRoot(xml: string): XmlNode {
	const body = xml.replace(/^\uFEFF?\s*<\?xml[^>]*\?>\s*/, "");
	assertWellFormed(body);
	const nodes = XmlNode.parse(body);
	const roots = nodes.filter((node) => !node.isText && node.tag !== "?xml");
	const root = roots[0];
	if (!root || roots.length !== 1) {
		throw new RawError("INVALID_XML", "A part has exactly one root element");
	}
	return root;
}
