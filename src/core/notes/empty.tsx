import { w } from "../jsx";
import type { XmlNode } from "../parser";
import { type NoteConfig, type NoteKind, noteConfig } from "./config";

/** Build an empty `<w:footnotes/>` or `<w:endnotes/>` tree seeded with the two
 *  reserved boilerplate entries: id=-1 separator (the rule above the note area)
 *  and id=0 continuationSeparator (used when notes wrap across pages). Both
 *  have `w:type` set; both are filtered out by `core/ast/read.ts`. We seed them
 *  so LibreOffice and Word render the note area correctly the first time an
 *  agent calls `footnotes add` against a doc that had none. */
export function buildEmptyNotesTree(kind: NoteKind): XmlNode[] {
	return [buildEmptyNotesRoot(noteConfig(kind))];
}

function buildEmptyNotesRoot(config: NoteConfig): XmlNode {
	const Root = config.kind === "footnote" ? w.footnotes : w.endnotes;
	return (
		<Root {...{ "xmlns:w": NS_W, "xmlns:r": NS_R }}>
			<NoteBoilerplate config={config} type="separator" id="-1" />
			<NoteBoilerplate config={config} type="continuationSeparator" id="0" />
		</Root>
	);
}

function NoteBoilerplate({
	config,
	type,
	id,
}: {
	config: NoteConfig;
	type: "separator" | "continuationSeparator";
	id: string;
}): XmlNode {
	const Note = config.kind === "footnote" ? w.footnote : w.endnote;
	return (
		<Note w-id={id} w-type={type}>
			<w.p>
				<w.r>
					{type === "separator" ? <w.separator /> : <w.continuationSeparator />}
				</w.r>
			</w.p>
		</Note>
	);
}

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
// `r:id` on note-body `<w:hyperlink>` (and any future note-body media) needs the
// relationships namespace declared on the part root, exactly as document.xml
// does — without it the part is malformed XML and Word reports "unreadable
// content". Declared unconditionally (Word always declares it on these parts).
const NS_R =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
