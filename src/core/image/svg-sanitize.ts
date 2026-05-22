// Strip active content from an SVG before we embed it into a docx. Other
// formats here are raster bytes; SVG is XML that downstream renderers (web
// previews, mail clients, anything later extracting and viewing the part) may
// execute. fast-xml-parser already rejects external-entity DOCTYPEs (XXE), so
// this layer handles the element/attribute level: scripts, event handlers,
// foreign-HTML islands, SSRF-on-render `href` targets, animation events.

import { XmlNode } from "../parser";

export function sanitizeSvg(bytes: Uint8Array): Uint8Array {
	const text = new TextDecoder().decode(bytes);
	const tree = XmlNode.parse(text);
	for (const node of tree) scrubNode(node);
	return new TextEncoder().encode(XmlNode.serialize(tree));
}

/** Elements we drop wholesale: they either execute (script, foreignObject's
 * HTML island) or open SSRF/event vectors (style with `@import`, link, the
 * animation family that supports `onbegin`/`onend`), or load foreign content
 * (audio/video/iframe/object/embed). */
const DROP_TAGS = new Set([
	"script",
	"foreignObject",
	"iframe",
	"object",
	"embed",
	"style",
	"link",
	"set",
	"animate",
	"animateMotion",
	"animateTransform",
	"audio",
	"video",
	"discard",
	"handler",
	"listener",
]);

const URL_ATTRS = ["href", "xlink:href"];

function scrubNode(node: XmlNode): void {
	node.children = node.children.filter(
		(child) => !DROP_TAGS.has(localName(child.tag)),
	);
	for (const key of Object.keys(node.attributes)) {
		// Drop every event handler attribute (`onclick`, `onload`, `onbegin`, …).
		if (/^on[a-z]/i.test(key)) delete node.attributes[key];
	}
	for (const attr of URL_ATTRS) {
		const value = node.attributes[attr];
		if (value !== undefined && !isAllowedUrl(value)) {
			delete node.attributes[attr];
		}
	}
	for (const child of node.children) scrubNode(child);
}

function localName(tag: string): string {
	const colon = tag.indexOf(":");
	return colon >= 0 ? tag.slice(colon + 1) : tag;
}

/** Allow only same-document fragments and inline RASTER `data:` payloads. Any
 * explicit scheme (`javascript:`, `data:` non-image or `data:image/svg+xml`,
 * `http(s):`, `file:`, …) is rejected; scheme-less relative refs are accepted
 * (they'd resolve against the SVG's location, which after embedding is the
 * docx — effectively inert). `data:image/svg+xml` is *not* allowed because the
 * inline SVG itself can carry script — e.g. `<use href="data:image/svg+xml,
 * <svg onload=evil()/>"/>` would bypass an `image/*` allowlist. */
function isAllowedUrl(url: string): boolean {
	const trimmed = url.trim().toLowerCase();
	if (trimmed.length === 0) return true;
	if (trimmed.startsWith("#")) return true;
	if (/^data:image\/(png|jpe?g|gif|webp|bmp)[;,]/.test(trimmed)) return true;
	if (/^[a-z][a-z0-9+.-]*:/.test(trimmed)) return false;
	return true;
}
