import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import JSZip from "jszip";

/**
 * Build tests/fixtures/notes.docx — a minimal docx exercising footnotes and
 * endnotes.
 *
 * Surfaces: w:footnoteReference, w:endnoteReference, footnotes.xml, endnotes.xml,
 * Word's reserved separator/continuationSeparator entries (filtered by
 * `readNotes` in core/ast/read.ts).
 */

const root = resolve(import.meta.dir, "..");
const out = resolve(root, "tests/fixtures/notes.docx");

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
	<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
	<Default Extension="xml" ContentType="application/xml"/>
	<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
	<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
	<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
	<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`;

const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
	<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
	<Relationship Id="rId21" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
</Relationships>`;

const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
	<w:body>
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
		<w:sectPr/>
	</w:body>
</w:document>`;

// Word writes two reserved entries before any user notes:
//   id="-1" w:type="separator" — the visual line above the footnote area
//   id="0"  w:type="continuationSeparator" — used when notes wrap across pages
// Both should be filtered out by core/ast/read.ts → readNotes (because they
// have w:type set). We include them here so the fixture exercises that path.
const footnotes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
	<w:footnote w:id="-1" w:type="separator">
		<w:p><w:r><w:separator/></w:r></w:p>
	</w:footnote>
	<w:footnote w:id="0" w:type="continuationSeparator">
		<w:p><w:r><w:continuationSeparator/></w:r></w:p>
	</w:footnote>
	<w:footnote w:id="1">
		<w:p>
			<w:r>
				<w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
				<w:footnoteRef/>
			</w:r>
			<w:r><w:t xml:space="preserve"> My note.</w:t></w:r>
		</w:p>
	</w:footnote>
</w:footnotes>`;

const endnotes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
	<w:endnote w:id="-1" w:type="separator">
		<w:p><w:r><w:separator/></w:r></w:p>
	</w:endnote>
	<w:endnote w:id="0" w:type="continuationSeparator">
		<w:p><w:r><w:continuationSeparator/></w:r></w:p>
	</w:endnote>
	<w:endnote w:id="1">
		<w:p>
			<w:r>
				<w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr>
				<w:endnoteRef/>
			</w:r>
			<w:r><w:t xml:space="preserve"> This is an endnote at the end of the document.</w:t></w:r>
		</w:p>
	</w:endnote>
</w:endnotes>`;

const coreProps = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">
	<dc:title>Notes Fixture</dc:title>
	<dc:creator>docx-cli</dc:creator>
	<dcterms:created>2026-05-05T00:00:00Z</dcterms:created>
	<dcterms:modified>2026-05-05T00:00:00Z</dcterms:modified>
</cp:coreProperties>`;

const zip = new JSZip();
zip.file("[Content_Types].xml", contentTypes);
zip.file("_rels/.rels", rootRels);
zip.file("word/document.xml", document);
zip.file("word/_rels/document.xml.rels", docRels);
zip.file("word/footnotes.xml", footnotes);
zip.file("word/endnotes.xml", endnotes);
zip.file("docProps/core.xml", coreProps);

const buf = await zip.generateAsync({
	type: "uint8array",
	compression: "DEFLATE",
});
mkdirSync(dirname(out), { recursive: true });
await Bun.write(out, buf);
console.log(`Wrote ${out} (${buf.length} bytes)`);
