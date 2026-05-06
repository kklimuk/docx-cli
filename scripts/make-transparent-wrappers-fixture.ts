import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import JSZip from "jszip";

/**
 * Build tests/fixtures/transparent-wrappers.docx — a minimal docx exercising
 * <w:fldSimple> (cached field results) and <w:smartTag> (semantic
 * annotations). Both wrap runs that contribute to paragraph text but were
 * historically invisible to the XML-side offset arithmetic in
 * comments/helpers.tsx, replace/replace-span.tsx, and hyperlinks/wrap.tsx.
 *
 * Layout:
 *   p0: "Today is " + <w:fldSimple w:instr="DATE">"2026-05-05"</w:fldSimple> + "."
 *   p1: "Hello " + <w:smartTag w:element="PersonName">"Alice"</w:smartTag> + "."
 */

const root = resolve(import.meta.dir, "..");
const out = resolve(root, "tests/fixtures/transparent-wrappers.docx");

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
	<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
	<Default Extension="xml" ContentType="application/xml"/>
	<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
	<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`;

const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
	<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;

const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
	<w:body>
		<w:p>
			<w:r><w:t xml:space="preserve">Today is </w:t></w:r>
			<w:fldSimple w:instr=" DATE \\@ &quot;yyyy-MM-dd&quot;">
				<w:r><w:t>2026-05-05</w:t></w:r>
			</w:fldSimple>
			<w:r><w:t>.</w:t></w:r>
		</w:p>
		<w:p>
			<w:r><w:t xml:space="preserve">Hello </w:t></w:r>
			<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="PersonName">
				<w:r><w:t>Alice</w:t></w:r>
			</w:smartTag>
			<w:r><w:t>.</w:t></w:r>
		</w:p>
		<w:sectPr/>
	</w:body>
</w:document>`;

const coreProps = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">
	<dc:title>Transparent Wrappers Fixture</dc:title>
	<dc:creator>docx-cli</dc:creator>
	<dcterms:created>2026-05-05T00:00:00Z</dcterms:created>
	<dcterms:modified>2026-05-05T00:00:00Z</dcterms:modified>
</cp:coreProperties>`;

const zip = new JSZip();
zip.file("[Content_Types].xml", contentTypes);
zip.file("_rels/.rels", rootRels);
zip.file("word/document.xml", document);
zip.file("word/_rels/document.xml.rels", docRels);
zip.file("docProps/core.xml", coreProps);

const buf = await zip.generateAsync({
	type: "uint8array",
	compression: "DEFLATE",
});
mkdirSync(dirname(out), { recursive: true });
await Bun.write(out, buf);
console.log(`Wrote ${out} (${buf.length} bytes)`);
