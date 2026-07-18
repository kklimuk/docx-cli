import type { XmlNode } from "./parser";

/** Markup-compatibility (ECMA-376 Part 3) helpers shared by everything that
 *  touches `mc:Ignorable` — the comments part (registering `w14`), the raw
 *  marker stamp (registering `dcx`), and the MCE preprocessing in the schema
 *  validator. One owner because the whitespace-split/append/rejoin idempotence
 *  is subtle enough to drift when copied. */
export const NS_MC =
	"http://schemas.openxmlformats.org/markup-compatibility/2006";

/** Register `prefix` in the root's `mc:Ignorable` list (declaring `xmlns:mc`
 *  first if absent). Idempotent. */
export function ensureIgnorable(root: XmlNode, prefix: string): void {
	if (!root.getAttribute("xmlns:mc")) root.setAttribute("xmlns:mc", NS_MC);
	const ignorable = (root.getAttribute("mc:Ignorable") ?? "")
		.split(/\s+/)
		.filter(Boolean);
	if (ignorable.includes(prefix)) return;
	ignorable.push(prefix);
	root.setAttribute("mc:Ignorable", ignorable.join(" "));
}

/** The namespace prefix of a `prefix:local` name, or undefined when bare. */
export function prefixOf(name: string): string | undefined {
	const colon = name.indexOf(":");
	return colon === -1 ? undefined : name.slice(0, colon);
}
