import { ensureIgnorable, NS_MC, prefixOf } from "./mc";
import type { XmlNode } from "./parser";

/** Declare, on a part's root element, every standard OOXML prefix the part USES
 *  but nobody declares. Our emitters write `r:id`, `wp:inline`, `m:oMath`,
 *  `w14:checkbox`, … assuming the root declares them the way Word's (and our
 *  `create` template's) roots do — but a docx from another producer may declare
 *  only `xmlns:w`, and an undeclared prefix is malformed XML: Word reports
 *  "unreadable content" and `validate` fails. Run on every part tree we
 *  serialize (fix, don't hint: the standard URIs are unambiguous). A Microsoft
 *  extension namespace (`w14`, `wp14`, …) also joins `mc:Ignorable`, as Word
 *  writes it, so a consumer that doesn't know it skips it. A prefix declared
 *  lower in the tree (`<a:graphic xmlns:a=…>`) counts as declared, and an
 *  unknown prefix is left alone — it isn't ours to guess. */
export function declareUsedNamespaces(tree: XmlNode[]): void {
	const root = tree.find(
		(node) => !node.tag.startsWith("?") && !node.tag.startsWith("#"),
	);
	if (!root) return;
	const missing = new Set<string>();
	collectUndeclaredPrefixes(root, new Set(), missing);
	for (const prefix of missing) declareKnownNamespace(root, prefix);
}

/** Declare a standard prefix on `root` (and register an extension prefix in
 *  `mc:Ignorable`). Returns false for a prefix outside the standard table, so
 *  the raw gate can reject it. */
export function declareKnownNamespace(root: XmlNode, prefix: string): boolean {
	const uri = KNOWN_OOXML_NAMESPACES[prefix];
	if (!uri) return false;
	if (!root.getAttribute(`xmlns:${prefix}`))
		root.setAttribute(`xmlns:${prefix}`, uri);
	if (IGNORABLE_EXTENSION_PREFIXES.has(prefix)) ensureIgnorable(root, prefix);
	return true;
}

function collectUndeclaredPrefixes(
	node: XmlNode,
	declaredAbove: Set<string>,
	missing: Set<string>,
): void {
	if (node.isText) return;
	let declared = declaredAbove;
	for (const key of Object.keys(node.attributes)) {
		if (!key.startsWith("xmlns:")) continue;
		if (declared === declaredAbove) declared = new Set(declaredAbove);
		declared.add(key.slice("xmlns:".length));
	}
	const used = [prefixOf(node.tag)];
	for (const key of Object.keys(node.attributes)) {
		if (!key.startsWith("xmlns")) used.push(prefixOf(key));
	}
	// `mc:Ignorable`/`mc:Choice Requires` NAME prefixes in their values too
	// ("w14 wp14"); a listed prefix must be declared or the part is invalid.
	const ignorable = node.attributes["mc:Ignorable"];
	if (ignorable) used.push(...ignorable.split(/\s+/));
	for (const prefix of used) {
		if (!prefix || prefix === "xml" || declared.has(prefix)) continue;
		missing.add(prefix);
	}
	for (const child of node.children) {
		collectUndeclaredPrefixes(child, declared, missing);
	}
}

/** The standard OOXML prefix → URI table (what Word declares on its own part
 *  roots). Shared with the raw gate, which auto-declares the same prefixes for
 *  a fragment that uses one undeclared. */
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
};

/** Extension namespaces Word lists in `mc:Ignorable` (they extend the ECMA
 *  schema; a consumer that predates them must skip, not reject). `wps`/`wpg`/…
 *  aren't here: they ride an `mc:Choice Requires=…` branch instead. */
const IGNORABLE_EXTENSION_PREFIXES = new Set([
	"w14",
	"w15",
	"w16",
	"w16se",
	"w16cid",
	"w16cex",
	"w16du",
	"w16sdtdh",
	"wp14",
]);
