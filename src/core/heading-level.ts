/** The outline level a paragraph style id names — `Heading1` → 1, `Heading 2`
 *  → 2, a bare `Heading` → 1 — or null when the style isn't a heading. The one
 *  rule `outline` builds its tree from and `move --subtree` expands a heading's
 *  section by, so the two can't disagree about where a section ends. */
export function headingLevel(
	style: string | undefined,
	stylePrefix = "Heading",
): number | null {
	if (!style) return null;
	if (!style.startsWith(stylePrefix)) return null;
	const remainder = style.slice(stylePrefix.length).trim();
	if (remainder === "") return 1;
	const parsed = Number(remainder);
	if (!Number.isInteger(parsed) || parsed < 1) return null;
	return parsed;
}
