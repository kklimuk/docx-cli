/** Footnotes and endnotes share their entire mechanics — separate parts, but
 *  identical XML shape, same id-allocation rules (Word reserves -1/0 for the
 *  separator/continuationSeparator boilerplate), same reference-run pattern,
 *  same body-paragraph shape. We parameterize on this kind everywhere instead
 *  of duplicating two near-identical modules. */
export function noteConfig(kind: NoteKind): NoteConfig {
	return kind === "footnote" ? FOOTNOTE_CONFIG : ENDNOTE_CONFIG;
}

export type NoteKind = "footnote" | "endnote";

export type NoteConfig = {
	kind: NoteKind;
	idPrefix: "fn" | "en";
	rootTag: "w:footnotes" | "w:endnotes";
	itemTag: "w:footnote" | "w:endnote";
	referenceTag: "w:footnoteReference" | "w:endnoteReference";
	bodyRefTag: "w:footnoteRef" | "w:endnoteRef";
	textStyle: "FootnoteText" | "EndnoteText";
	referenceStyle: "FootnoteReference" | "EndnoteReference";
	partName: "word/footnotes.xml" | "word/endnotes.xml";
	target: "footnotes.xml" | "endnotes.xml";
	relationshipType: string;
	contentType: string;
};

const FOOTNOTES_RELATIONSHIP_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
const ENDNOTES_RELATIONSHIP_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
const FOOTNOTES_CONTENT_TYPE =
	"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
const ENDNOTES_CONTENT_TYPE =
	"application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";

const FOOTNOTE_CONFIG: NoteConfig = {
	kind: "footnote",
	idPrefix: "fn",
	rootTag: "w:footnotes",
	itemTag: "w:footnote",
	referenceTag: "w:footnoteReference",
	bodyRefTag: "w:footnoteRef",
	textStyle: "FootnoteText",
	referenceStyle: "FootnoteReference",
	partName: "word/footnotes.xml",
	target: "footnotes.xml",
	relationshipType: FOOTNOTES_RELATIONSHIP_TYPE,
	contentType: FOOTNOTES_CONTENT_TYPE,
};

const ENDNOTE_CONFIG: NoteConfig = {
	kind: "endnote",
	idPrefix: "en",
	rootTag: "w:endnotes",
	itemTag: "w:endnote",
	referenceTag: "w:endnoteReference",
	bodyRefTag: "w:endnoteRef",
	textStyle: "EndnoteText",
	referenceStyle: "EndnoteReference",
	partName: "word/endnotes.xml",
	target: "endnotes.xml",
	relationshipType: ENDNOTES_RELATIONSHIP_TYPE,
	contentType: ENDNOTES_CONTENT_TYPE,
};
