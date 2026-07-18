import type { Document } from "../ast/document";
import { ensureIgnorable, NS_MC, prefixOf } from "../mc";
import { XmlNode } from "../parser";
import { RawError } from "./error";

/** Gate 4: every prefix a fragment uses must resolve — either declared inside
 *  the fragment itself or on the document root. A missing KNOWN prefix is
 *  auto-declared on `<w:document>` (fix, don't hint: the standard URIs are
 *  unambiguous and doc-root declaration is where Word itself puts them); an
 *  unknown prefix rejects, because Word reports an undeclared prefix as
 *  "unreadable content." */
export function ensureFragmentNamespaces(
	nodes: XmlNode[],
	document: Document,
): void {
	const documentRoot = requireDocumentRoot(document);
	for (const node of nodes) {
		checkPrefixes(node, documentRoot, new Set());
	}
}

/** Stamp each inserted root with `dcx:raw="1"` and register our `dcx` prefix
 *  in `mc:Ignorable` on the doc root. Word (and the validator, verified)
 *  ignores attributes in an Ignorable namespace, while our reader surfaces the
 *  marker as the `docx:raw` annotation — so a raw insertion stays identifiable
 *  across the write-read loop. A resave by Word/LibreOffice may drop the
 *  marker; the content stays, only the annotation degrades. */
export function stampRawMarkers(nodes: XmlNode[], document: Document): void {
	const documentRoot = requireDocumentRoot(document);
	declareNamespace(documentRoot, RAW_MARKER_PREFIX, RAW_MARKER_NAMESPACE);
	ensureIgnorable(documentRoot, RAW_MARKER_PREFIX);
	for (const node of nodes) {
		node.setAttribute(RAW_MARKER_ATTRIBUTE, "1");
	}
}

const RAW_MARKER_PREFIX = "dcx";
const RAW_MARKER_NAMESPACE = "urn:docx-cli:raw";
export const RAW_MARKER_ATTRIBUTE = `${RAW_MARKER_PREFIX}:raw`;

export function requireDocumentRoot(document: Document): XmlNode {
	const root = XmlNode.findRoot(document.documentTree, "w:document");
	if (!root) {
		throw new RawError("USAGE", "Document has no <w:document> root");
	}
	return root;
}

function declareNamespace(root: XmlNode, prefix: string, uri: string): void {
	if (!root.getAttribute(`xmlns:${prefix}`)) {
		root.setAttribute(`xmlns:${prefix}`, uri);
	}
}

function checkPrefixes(
	node: XmlNode,
	documentRoot: XmlNode,
	declaredAbove: Set<string>,
): void {
	if (node.isText) return;
	const declaredHere = new Set(declaredAbove);
	for (const key of Object.keys(node.attributes)) {
		if (key.startsWith("xmlns:")) declaredHere.add(key.slice("xmlns:".length));
	}
	const used = [prefixOf(node.tag)];
	for (const key of Object.keys(node.attributes)) {
		if (!key.startsWith("xmlns")) used.push(prefixOf(key));
	}
	for (const prefix of used) {
		if (!prefix || prefix === "xml") continue;
		if (declaredHere.has(prefix)) continue;
		if (documentRoot.getAttribute(`xmlns:${prefix}`)) continue;
		const uri = KNOWN_OOXML_NAMESPACES[prefix];
		if (uri) {
			documentRoot.setAttribute(`xmlns:${prefix}`, uri);
			continue;
		}
		throw new RawError(
			"INVALID_XML",
			`Fragment uses undeclared namespace prefix "${prefix}:"`,
			`Declare it on your fragment root (xmlns:${prefix}="…") — an undeclared prefix makes Word report the file as unreadable.`,
		);
	}
	for (const child of node.children) {
		checkPrefixes(child, documentRoot, declaredHere);
	}
}

/** The standard OOXML prefix → URI table (what Word declares on its own doc
 *  roots). A fragment using one of these without a declaration gets it added
 *  to `<w:document>` automatically. */
const KNOWN_OOXML_NAMESPACES: Record<string, string> = {
	w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
	r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
	m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
	mc: NS_MC,
	a: "http://schemas.openxmlformats.org/drawingml/2006/main",
	pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
	wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
	c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
	dgm: "http://schemas.openxmlformats.org/drawingml/2006/diagram",
	v: "urn:schemas-microsoft-com:vml",
	o: "urn:schemas-microsoft-com:office:office",
	w10: "urn:schemas-microsoft-com:office:word",
	w14: "http://schemas.microsoft.com/office/word/2010/wordml",
	w15: "http://schemas.microsoft.com/office/word/2012/wordml",
	w16: "http://schemas.microsoft.com/office/word/2018/wordml",
	w16se: "http://schemas.microsoft.com/office/word/2015/wordml/symex",
	w16cid: "http://schemas.microsoft.com/office/word/2016/wordml/cid",
	w16cex: "http://schemas.microsoft.com/office/word/2018/wordml/cex",
	w16du: "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
	w16sdtdh: "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
	wp14: "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
	wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
	wpg: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
	wpc: "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
	wpi: "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
	cx: "http://schemas.microsoft.com/office/drawing/2014/chartex",
	[RAW_MARKER_PREFIX]: RAW_MARKER_NAMESPACE,
};
