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

/** The branch of an `<mc:AlternateContent>` a consumer reads: its FIRST
 *  `<mc:Choice>` (what Word itself takes — the spec requires at least one and
 *  the consumer picks the first it understands), else the `<mc:Fallback>` for
 *  a lenient read of a producer that wrote only that. Undefined for an empty
 *  wrapper. MCE lets a producer wrap ANY element at ANY level — a block, a
 *  paragraph's runs, a run's children — and nest wrappers inside branches, so
 *  every walker resolves through this one helper rather than special-casing
 *  the one shape Word writes (a `wps` shape inside `<w:r>`). */
export function alternateContentBranch(node: XmlNode): XmlNode | undefined {
	return node.findChild("mc:Choice") ?? node.findChild("mc:Fallback");
}

export function isAlternateContent(node: XmlNode): boolean {
	return node.tag === "mc:AlternateContent";
}

/** Every `<w:txbxContent>` story reachable from `node`, in document order,
 *  resolving nested `<mc:AlternateContent>` to ONE branch as it descends so
 *  the Choice copy and its Fallback twin never both surface. A `wpg` group
 *  shape yields one story per text box it holds. */
export function collectTextBoxContents(node: XmlNode): XmlNode[] {
	const out: XmlNode[] = [];
	if (isAlternateContent(node)) {
		const branch = alternateContentBranch(node);
		return branch ? collectTextBoxContents(branch) : out;
	}
	const visit = (current: XmlNode): void => {
		for (const child of current.children) {
			if (child.tag === "w:txbxContent") {
				out.push(child);
				continue;
			}
			if (isAlternateContent(child)) {
				const branch = alternateContentBranch(child);
				if (branch) visit(branch);
				continue;
			}
			visit(child);
		}
	};
	visit(node);
	return out;
}

/** Re-copy every text box's Choice story into its Fallback twin before a save.
 *
 *  Word writes each text box TWICE — the `wps` shape under `<mc:Choice>` and a
 *  VML `<w:pict>` under `<mc:Fallback>`, both carrying the same
 *  `<w:txbxContent>` — and keeps the pair in sync on every save. Our reader
 *  addresses the Choice copy (`tbxN:pK` locators point into it), so after an
 *  edit the Fallback would show the OLD text to any consumer that reads it
 *  (Word 2007, some converters). Mirroring the Choice story over it on save is
 *  exactly what Word does; a wrapper whose branches don't pair up story-for-
 *  story is left alone. Returns the number of stories synced. */
export function syncTextBoxFallbacks(tree: XmlNode[]): number {
	let synced = 0;
	const visit = (node: XmlNode): void => {
		if (isAlternateContent(node)) {
			const choice = node.findChild("mc:Choice");
			const fallback = node.findChild("mc:Fallback");
			// POST-order: a box nested inside this one syncs its own pair first,
			// so the outer clone below carries the inner box's fresh Fallback.
			if (choice) visit(choice);
			if (choice && fallback) {
				const sources = collectTextBoxContents(choice);
				const targets = collectTextBoxContents(fallback);
				if (sources.length > 0 && sources.length === targets.length) {
					for (const [index, source] of sources.entries()) {
						const target = targets[index];
						if (!target) continue;
						target.children = source.children.map((child) => child.clone());
						synced += 1;
					}
				}
			}
			// The fallback was just mirrored from the choice (or is unpaired), so
			// it never recurses.
			return;
		}
		for (const child of node.children) visit(child);
	};
	for (const root of tree) visit(root);
	return synced;
}
