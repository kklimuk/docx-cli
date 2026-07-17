import { XMLValidator } from "fast-xml-parser";
import { XmlNode } from "../parser";
import { RawError } from "./error";

/** Gates 1–2.5: fragment parsing. Well-formedness, the addressable-roots rule,
 *  and the inter-element whitespace strip — the pieces every raw surface
 *  (blocks, relationships, parts) runs before anything domain-specific. */

/** What roots a fragment may have. `blocks`: `w:p`/`w:tbl` (insert, and
 *  replace at a pN/tN). `sectPr`: exactly one `w:sectPr` (replace at an sN).
 *  `relationships`: `<Relationship>` roots only (the rels part). */
export type FragmentRoots = "blocks" | "sectPr" | "relationships";

/** Gate 1. `XMLValidator` is the only rejecting parser in the dependency tree —
 *  `XmlNode.parse` recovers silently from anything. After well-formedness is
 *  proven, a literal `<!DOCTYPE`/`<!--`/`<![CDATA[`/`<?` substring can only be
 *  the real construct (inside text it would be entity-escaped), and all four
 *  are things our parser drops or rewrites rather than preserves — reject them
 *  so nothing silently disappears between the argv and the file. */
export function assertWellFormed(xml: string): void {
	// A fragment may legally have several roots (<w:p/><w:p/>), which
	// XMLValidator rejects as a document — validate inside a synthetic wrapper
	// and shift first-line columns back by its length. The wrapper's name is
	// chosen to read sensibly when it leaks into an unclosed-tag message
	// ("expected </w:t> instead of closing tag 'end-of-fragment'").
	const wrapper = "<end-of-fragment>";
	const verdict = XMLValidator.validate(`${wrapper}${xml}</end-of-fragment>`);
	if (verdict !== true) {
		const { msg, line, col } = verdict.err;
		const column = line === 1 ? Math.max(1, col - wrapper.length) : col;
		throw new RawError(
			"INVALID_XML",
			`Fragment is not well-formed XML (line ${line}, col ${column}): ${msg}`,
		);
	}
	const rejects: [string, string][] = [
		["<!DOCTYPE", "a DOCTYPE has no place inside a document part"],
		["<!--", "XML comments are dropped by the parser — remove them"],
		[
			"<![CDATA[",
			"CDATA sections are rewritten by the parser — inline the text with entities (&lt; &amp;) instead",
		],
		[
			"<?",
			"processing instructions / XML declarations are dropped — pass the fragment without them",
		],
	];
	const lowered = xml.toLowerCase();
	for (const [needle, reason] of rejects) {
		if (lowered.includes(needle.toLowerCase())) {
			throw new RawError(
				"INVALID_XML",
				`Fragment contains ${needle} — ${reason}`,
			);
		}
	}
}

/** Gate 2: parse and enforce the addressable-roots rule. Everything INSIDE a
 *  root is unrestricted — that's the escape-hatch value — but the roots
 *  themselves must be elements the reader assigns locators to, or the insert
 *  would be invisible to the next `read` and unaddressable forever after. */
export function parseRoots(xml: string, allow: FragmentRoots): XmlNode[] {
	// Parse inside a wrapper element: the parser silently DROPS text at the
	// top level of a document, but preserves it as mixed content inside an
	// element — so wrapping is what lets the stray-text check below see it.
	const wrapped = XmlNode.parse(`<x>${xml}</x>`);
	const roots = (wrapped[0]?.children ?? []).filter((node) => {
		if (!node.isText) return true;
		if ((node.text ?? "").trim() === "") return false;
		throw new RawError(
			"INVALID_XML",
			"Fragment has text outside any element — wrap it in a <w:p><w:r><w:t>…</w:t></w:r></w:p>",
		);
	});
	if (roots.length === 0) {
		throw new RawError("INVALID_XML", "Fragment is empty");
	}
	if (allow === "relationships") {
		for (const root of roots) {
			if (root.tag === "Relationship") continue;
			throw new RawError(
				"INVALID_XML",
				`Relationship fragments take only <Relationship> roots — got <${root.tag}>`,
				"One destination per call: body blocks (w:p/w:tbl) and relationships can't mix in one fragment.",
			);
		}
		return roots;
	}
	if (allow === "sectPr") {
		const only = roots[0];
		if (roots.length !== 1 || !only || only.tag !== "w:sectPr") {
			throw new RawError(
				"INVALID_XML",
				"Replacing a section (sN) takes exactly one <w:sectPr> root",
			);
		}
		return roots;
	}
	for (const root of roots) {
		if (root.tag === "w:p" || root.tag === "w:tbl") continue;
		throw new RawError(
			"INVALID_XML",
			`Fragment root <${root.tag}> is not addressable — top-level elements must be <w:p> or <w:tbl>`,
			rootHintFor(root.tag),
		);
	}
	return roots;
}

function rootHintFor(tag: string): string {
	if (tag === "w:sectPr") {
		return "Sections are addressed as sN — use `raw replace --at sN`, or the `docx sections` verbs.";
	}
	if (tag === "Relationship") {
		return "Relationship fragments route by their root and are single-shot only (rIds don't shift, so no batch is needed): run one `raw insert FILE --xml '<Relationship …/>'` per call, without placement flags, and don't mix them with block roots.";
	}
	return "Only <w:p>/<w:tbl> get locators (pN/tN), so anything else would vanish from `docx read` and could never be addressed again. Wrap the construct inside a <w:p> or <w:tbl>.";
}

/** Gate 2.5. The parser keeps whitespace (`trimValues: false`), so a fragment
 *  authored with indentation lands `#text` children inside container elements —
 *  and whitespace character data inside a complex-content parent (`w:p`,
 *  `w:tbl`, `w:pPr`, …) makes Word repair the file. Strip whitespace-only text
 *  children from any element that has at least one element child; leaf text
 *  (`w:t`, which may carry `xml:space="preserve"`) is untouched. */
export function stripInterElementWhitespace(node: XmlNode): void {
	const hasElementChild = node.children.some((child) => !child.isText);
	if (hasElementChild) {
		node.children = node.children.filter(
			(child) => !(child.isText && (child.text ?? "").trim() === ""),
		);
	}
	for (const child of node.children) {
		if (!child.isText) stripInterElementWhitespace(child);
	}
}
