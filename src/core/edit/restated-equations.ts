import { latexToOmml } from "../equation/emit";
import { equationToken } from "../equation/token";
import { isEquation, type XmlNode } from "../parser";
import {
	type OldEquation,
	visibleEquations,
} from "../track-changes/equation-text";

/** Put back each equation a whole-paragraph markdown replace RE-STATES, and
 *  return the old equations it doesn't (the ones the edit deletes).
 *
 *  The markdown importer compiles every `$…$` afresh, so even a token copied
 *  verbatim from `read` comes back as a NEW `<m:oMath>` — LaTeX can't carry
 *  everything the original OMML held (Word's run formatting, `m:ctrlPr`,
 *  constructs the reader degrades). So a new equation whose token equals an
 *  old one's (matched in reading order, each old one at most once) is swapped
 *  for the ORIGINAL node — a clone under tracking, since the original rides
 *  the `<w:del>` of the replaced content. Comparing the token as compiled
 *  too absorbs LaTeX the writer normalizes (`x^2` ↔ `x^{2}`). */
export function keepRestatedEquations(
	oldParagraphs: XmlNode[],
	newBlocks: XmlNode[],
	tracked: boolean,
): OldEquation[] {
	const old = oldParagraphs.flatMap((paragraph) => visibleEquations(paragraph));
	if (old.length === 0) return [];
	const placed = newBlocks.flatMap((block) => equationSlots(block));
	const claimed = new Set<number>();
	const dropped: OldEquation[] = [];
	for (const equation of old) {
		const spellings = tokenSpellings(equation);
		const slotIndex = placed.findIndex(
			(slot, index) => !claimed.has(index) && spellings.has(slot.token),
		);
		const slot = placed[slotIndex];
		if (!slot) {
			dropped.push(equation);
			continue;
		}
		claimed.add(slotIndex);
		const position = slot.siblings.indexOf(slot.node);
		slot.siblings.splice(
			position,
			1,
			tracked ? equation.node.clone() : equation.node,
		);
	}
	return dropped;
}

/** Every equation under `node`, with the live child array holding it. */
function equationSlots(
	node: XmlNode,
): { node: XmlNode; siblings: XmlNode[]; token: string }[] {
	const out: { node: XmlNode; siblings: XmlNode[]; token: string }[] = [];
	for (const child of node.children) {
		if (child.isText) continue;
		if (isEquation(child)) {
			const token = equationToken(child);
			if (token !== null)
				out.push({ node: child, siblings: node.children, token });
			continue;
		}
		out.push(...equationSlots(child));
	}
	return out;
}

/** The token as `read` shows it, and as the writer would compile it back. */
function tokenSpellings(equation: OldEquation): Set<string> {
	const spellings = new Set([equation.token]);
	const latex = equation.token.replace(/^\$+|\$+$/g, "");
	try {
		const recompiled = equationToken(latexToOmml(latex, equation.display));
		if (recompiled) spellings.add(recompiled);
	} catch {
		// LaTeX the writer can't compile can't have been re-stated either.
	}
	return spellings;
}
