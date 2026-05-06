import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import JSZip from "jszip";

/**
 * Build tests/fixtures/tracked-moves.docx — a minimal docx exercising
 * <w:moveFrom> / <w:moveTo> tracked moves. Models a realistic scenario:
 * one sentence was moved from p0 (origin) to p1 (destination).
 *
 * Layout:
 *   p0: "Origin paragraph: " + <w:moveFrom>"the moved sentence"</w:moveFrom> + "."
 *   p1: "Destination paragraph: " + <w:moveTo>"the moved sentence"</w:moveTo> + "."
 *
 * Views:
 *   current   = both occurrences visible (what's on disk)
 *   accepted  = only the destination retains the sentence; origin drops it
 *   baseline  = only the origin has the sentence; destination is empty
 *
 * moveFrom uses <w:t> (matching Word's typical output) rather than <w:delText>;
 * our reject codepath handles both via a no-op-for-<w:t> renameDelTextToText.
 */

const root = resolve(import.meta.dir, "..");
const out = resolve(root, "tests/fixtures/tracked-moves.docx");

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
			<w:r><w:t xml:space="preserve">Origin paragraph: </w:t></w:r>
			<w:moveFrom w:id="1" w:author="Reviewer" w:date="2026-05-05T12:00:00Z">
				<w:r><w:t>the moved sentence</w:t></w:r>
			</w:moveFrom>
			<w:r><w:t>.</w:t></w:r>
		</w:p>
		<w:p>
			<w:r><w:t xml:space="preserve">Destination paragraph: </w:t></w:r>
			<w:moveTo w:id="2" w:author="Reviewer" w:date="2026-05-05T12:00:00Z">
				<w:r><w:t>the moved sentence</w:t></w:r>
			</w:moveTo>
			<w:r><w:t>.</w:t></w:r>
		</w:p>
		<w:sectPr/>
	</w:body>
</w:document>`;

const coreProps = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">
	<dc:title>Tracked Moves Fixture</dc:title>
	<dc:creator>docx-cli</dc:creator>
	<dcterms:created>2026-05-05T00:00:00Z</dcterms:created>
	<dcterms:modified>2026-05-05T12:00:00Z</dcterms:modified>
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
