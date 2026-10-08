import { writeStderr } from "../respond";

/** Warn that a content edit resolved other authors' pending tracked changes
 *  without review (issue #17). The edit still succeeds — this is the interim
 *  behavior until the rebuild keeps them — so the warning names the span-edit
 *  alternatives that DO keep them pending. `paragraphLocator` is the paragraph
 *  the `:START-END` example addresses: the edited locator by default, a bare
 *  cell's sole `tN:rRcC:p0` (a cell locator takes no span), or null for a
 *  multi-paragraph range. `tracked` carries the edit's own tracking into both
 *  suggestions, so following one records the change the same way. */
export async function warnPendingRevisions(
	locator: string,
	authors: string[],
	{
		paragraphLocator = locator,
		tracked,
	}: { paragraphLocator?: string | null; tracked: boolean },
): Promise<void> {
	if (authors.length === 0) return;
	const track = tracked ? " --track" : "";
	await writeStderr(
		`warning: ${locator} had pending tracked changes by ${authors.join(", ")}; ` +
			"replacing whole paragraphs resolved them without review (they can no longer be accepted or rejected as theirs). " +
			`To keep them pending, edit only the words that change: \`docx replace FILE OLD NEW${track}\` or \`docx edit FILE --at ${paragraphLocator ?? "pN"}:START-END --text NEW${track}\`.\n`,
	);
}
