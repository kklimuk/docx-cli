import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import JSZip from "jszip";
import {
	addCanonicalParts,
	buildContentTypes,
	buildCoreProps,
	buildDocumentRels,
	buildRootRels,
	type ExtraPart,
	wrapDocument,
} from "./fixture-helpers";

/**
 * Build tests/fixtures/notes.docx — a minimal docx exercising footnotes and
 * endnotes.
 *
 * Surfaces: w:footnoteReference, w:endnoteReference, footnotes.xml,
 * endnotes.xml, Word's reserved separator/continuationSeparator entries
 * (filtered by `readNotes` in core/ast/read.ts).
 */

const root = resolve(import.meta.dir, "..");
const out = resolve(root, "tests/fixtures/notes.docx");

const body = `
	<w:p>
		<w:r><w:t xml:space="preserve">Test footnote.</w:t></w:r>
		<w:r>
			<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
			<w:footnoteReference w:id="1"/>
		</w:r>
		<w:r><w:t xml:space="preserve"> Test endnote.</w:t></w:r>
		<w:r>
			<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
			<w:endnoteReference w:id="1"/>
		</w:r>
	</w:p>
	<w:sectPr/>`;

// Word writes two reserved entries before any user notes:
//   id="-1" w:type="separator" — the visual line above the footnote area
//   id="0"  w:type="continuationSeparator" — used when notes wrap across pages
// Both should be filtered out by core/ast/read.ts → readNotes (because they
// have w:type set). We include them here so the fixture exercises that path.
const footnotes: ExtraPart = {
	partName: "word/footnotes.xml",
	contentType:
		"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
	relationshipType:
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
	target: "footnotes.xml",
	body: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="0" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> My note.</w:t></w:r></w:p></w:footnote></w:footnotes>`,
};

const endnotes: ExtraPart = {
	partName: "word/endnotes.xml",
	contentType:
		"application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
	relationshipType:
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
	target: "endnotes.xml",
	body: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:id="0" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="1"><w:p><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> This is an endnote at the end of the document.</w:t></w:r></w:p></w:endnote></w:endnotes>`,
};

const extras: ExtraPart[] = [footnotes, endnotes];

const zip = new JSZip();
zip.file("[Content_Types].xml", buildContentTypes(extras));
zip.file("_rels/.rels", buildRootRels());
zip.file("word/document.xml", wrapDocument(body));
zip.file("word/_rels/document.xml.rels", buildDocumentRels(extras));
zip.file(
	"docProps/core.xml",
	buildCoreProps({
		title: "Notes Fixture",
		author: "docx-cli",
		created: "2026-05-05T00:00:00Z",
	}),
);
for (const part of extras) zip.file(part.partName, part.body);
addCanonicalParts(zip);

const buf = await zip.generateAsync({
	type: "uint8array",
	compression: "DEFLATE",
});
mkdirSync(dirname(out), { recursive: true });
await Bun.write(out, buf);
console.log(`Wrote ${out} (${buf.length} bytes)`);
