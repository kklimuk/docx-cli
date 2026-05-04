import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import JSZip from "jszip";

const root = resolve(import.meta.dir, "..");
const out = resolve(root, "tests/fixtures/minimal.docx");

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
	<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
	<Default Extension="xml" ContentType="application/xml"/>
	<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
	<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
	<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`;

const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
	<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>`;

const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
	<w:body>
		<w:p>
			<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
			<w:r><w:t>Style Guide</w:t></w:r>
		</w:p>
		<w:p>
			<w:r><w:t xml:space="preserve">Use </w:t></w:r>
			<w:r>
				<w:rPr><w:color w:val="800080"/><w:b/></w:rPr>
				<w:t>important</w:t>
			</w:r>
			<w:r><w:t xml:space="preserve"> terms in purple bold.</w:t></w:r>
		</w:p>
		<w:p>
			<w:r><w:t xml:space="preserve">The quick brown </w:t></w:r>
			<w:commentRangeStart w:id="0"/>
			<w:r><w:t>fox</w:t></w:r>
			<w:commentRangeEnd w:id="0"/>
			<w:r>
				<w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
				<w:commentReference w:id="0"/>
			</w:r>
			<w:r><w:t xml:space="preserve"> jumps over the lazy dog.</w:t></w:r>
		</w:p>
		<w:sectPr/>
	</w:body>
</w:document>`;

const comments = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
	<w:comment w:id="0" w:author="Jane" w:date="2026-04-30T08:00:00Z" w:initials="J">
		<w:p><w:r><w:t>Should this be 'cat'?</w:t></w:r></w:p>
	</w:comment>
</w:comments>`;

const coreProps = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">
	<dc:title>Test Doc</dc:title>
	<dc:creator>Kirill</dc:creator>
	<dcterms:created>2026-05-01T00:00:00Z</dcterms:created>
	<dcterms:modified>2026-05-03T00:00:00Z</dcterms:modified>
</cp:coreProperties>`;

const zip = new JSZip();
zip.file("[Content_Types].xml", contentTypes);
zip.file("_rels/.rels", rootRels);
zip.file("word/document.xml", document);
zip.file("word/_rels/document.xml.rels", docRels);
zip.file("word/comments.xml", comments);
zip.file("docProps/core.xml", coreProps);

const buf = await zip.generateAsync({
	type: "uint8array",
	compression: "DEFLATE",
});
mkdirSync(dirname(out), { recursive: true });
await Bun.write(out, buf);
console.log(`Wrote ${out} (${buf.length} bytes)`);
