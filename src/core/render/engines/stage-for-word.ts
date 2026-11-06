import { Pkg } from "../../ast/document/package";

/** Copy `inputDocx` to `stagedDocx` for Word to open — minus the markup that
 *  makes Word BLOCK on open with a question nobody is there to answer.
 *
 *  Word asks "This document contains fields that may refer to other files. Do
 *  you want to update the fields in this document?" when it opens a document
 *  whose fields are flagged stale: a `w:dirty="true"` on a field's
 *  `<w:fldChar w:fldCharType="begin">` / `<w:fldSimple>` (verified: such a file
 *  popped the modal, its dirty-free twin rendered cleanly), or the document-level
 *  `<w:updateFields w:val="true"/>` in `settings.xml` (Word writes it after
 *  inserting a TOC). `set display alerts to none` does not suppress the prompt —
 *  it sat on the user's screen while renders waited on it, Word bouncing in the
 *  Dock — so an unattended render hung until a human clicked.
 *
 *  Stripping both from the STAGED copy renders exactly what clicking "No" shows
 *  (the field results saved in the file) and never touches the user's document.
 *  Only parts that carry the markup are rewritten — as text, so nothing else in
 *  them is re-serialized; every other part is copied as stored. Word for Windows
 *  opens the same staged copy, so a prompt there can't wedge the COM call either.
 *
 *  An input that isn't a readable zip is copied as-is: Word's own corrupt-file
 *  handling (and the honest RENDER_FAILED it produces) stays unchanged. */
export async function stageForWord(
	inputDocx: string,
	stagedDocx: string,
): Promise<void> {
	const pkg = await openOrNull(inputDocx);
	const changed = pkg ? await stripUpdatePrompts(pkg) : false;
	if (!pkg || !changed) {
		await Bun.write(stagedDocx, Bun.file(inputDocx));
		return;
	}
	await pkg.save(stagedDocx);
}

async function openOrNull(path: string): Promise<Pkg | null> {
	try {
		return await Pkg.open(path);
	} catch {
		return null;
	}
}

/** Remove the update-fields-on-open flag and every dirty field marker from the
 *  parts that can hold them. Returns whether anything changed. */
async function stripUpdatePrompts(pkg: Pkg): Promise<boolean> {
	let changed = false;
	for (const part of pkg.listParts()) {
		if (!PROMPTING_PART.test(part)) continue;
		const xml = await pkg.readText(part);
		const stripped = xml.replace(UPDATE_FIELDS, "").replace(DIRTY_FIELD, "$1");
		if (stripped === xml) continue;
		pkg.writeText(part, stripped);
		changed = true;
	}
	return changed;
}

/** The parts a field (or the settings flag) can live in: the body, headers,
 *  footers, notes, comments, and settings. */
const PROMPTING_PART =
	/^word\/(?:document|settings|footnotes|endnotes|comments|header\d*|footer\d*)\.xml$/;

/** `<w:updateFields …/>` (any prefix; self-closing, or an empty open/close
 *  pair). CT_OnOff has no children, so a text-level strip is exact. */
const UPDATE_FIELDS =
	/<(?:[A-Za-z_][\w.-]*:)?updateFields\b[^>]*?(?:\/>|>\s*<\/(?:[A-Za-z_][\w.-]*:)?updateFields>)/g;

/** A `dirty` attribute on a field marker (`<w:fldChar … w:dirty="true"/>`,
 *  `<w:fldSimple … w:dirty="1">`), captured up to the attribute so the replace
 *  keeps the rest of the tag. `dirty` exists only on these two elements. */
const DIRTY_FIELD =
	/(<(?:[A-Za-z_][\w.-]*:)?(?:fldChar|fldSimple)\b[^>]*?)\s+(?:[A-Za-z_][\w.-]*:)?dirty\s*=\s*(?:"[^"]*"|'[^']*')/g;
