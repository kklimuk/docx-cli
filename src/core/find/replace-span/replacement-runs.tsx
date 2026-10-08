import { textToRunElements } from "../../blocks";
import { applyRunFormatToRpr } from "../../edit/set-formatting";
import { w } from "../../jsx";
import { XmlNode } from "../../parser";
import type { ReplacementFormatting } from ".";

/** The replacement's run(s). Routed through `textToRunElements` so a real tab
 *  becomes `<w:tab/>` instead of a raw control character inside `<w:t>` — the
 *  same real-character handling every inline authoring surface uses. Plain
 *  single-line text still collapses to one `<w:t>` run, byte-identical to the
 *  old single-run shape; each produced run inherits the span's rPr. Returned
 *  as a plain array (never a fragment sentinel) so the nodes land directly in
 *  the live tree — `document.reread()` between batch entries must see them.
 *  (Real newlines never reach here — the CLI routes any `\n`-bearing
 *  replacement to the cross-paragraph path, where it means a paragraph mark.)
 *  Shared with replace-across.tsx, which builds each segment's runs the same
 *  way. */
export function replacementRuns(
	runProperties: XmlNode | null,
	text: string,
	formatting?: ReplacementFormatting,
): XmlNode[] {
	const properties = replacementRunProperties(runProperties, formatting);
	if (text.length === 0) {
		return [
			<w.r>
				{properties}
				<w.t {...{ "xml:space": "preserve" }} />
			</w.r>,
		];
	}
	const runs = textToRunElements(text);
	for (const run of runs) {
		if (properties) run.children.unshift(properties.clone());
	}
	return runs;
}

function replacementRunProperties(
	runProperties: XmlNode | null,
	formatting?: ReplacementFormatting,
): XmlNode | null {
	let properties = runProperties?.clone() ?? null;
	if (properties && formatting?.clearTags) {
		properties.children = properties.children.filter(
			(child) => !formatting.clearTags?.has(child.tag),
		);
		if (properties.children.length === 0) properties = null;
	}
	if (formatting?.format) {
		properties ??= XmlNode.element("w:rPr");
		applyRunFormatToRpr(properties, formatting.format);
	}
	return properties;
}
