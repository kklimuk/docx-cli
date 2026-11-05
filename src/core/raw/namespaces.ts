import type { Document } from "../ast/document";
import { ensureIgnorable, prefixOf } from "../mc";
import { declareKnownNamespace } from "../namespaces";
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
		if (prefix === RAW_MARKER_PREFIX) {
			declareNamespace(documentRoot, RAW_MARKER_PREFIX, RAW_MARKER_NAMESPACE);
			continue;
		}
		if (declareKnownNamespace(documentRoot, prefix)) continue;
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
