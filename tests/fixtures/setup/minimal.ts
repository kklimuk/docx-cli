import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import JSZip from "jszip";
import {
	addCanonicalParts,
	buildContentTypes,
	buildCoreProps,
	buildDocumentRels,
	buildRootRels,
	DEFAULT_SECTPR,
	type ExtraPart,
	wrapDocument,
} from "./helpers";

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/minimal.docx");

const body = `
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
	${DEFAULT_SECTPR}`;

const comments: ExtraPart = {
	partName: "word/comments.xml",
	contentType:
		"application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
	relationshipType:
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
	target: "comments.xml",
	body: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="Jane" w:date="2026-04-30T08:00:00Z" w:initials="J"><w:p><w:r><w:t>Should this be 'cat'?</w:t></w:r></w:p></w:comment></w:comments>`,
};

const extras: ExtraPart[] = [comments];

const zip = new JSZip();
zip.file("[Content_Types].xml", buildContentTypes(extras));
zip.file("_rels/.rels", buildRootRels());
zip.file("word/document.xml", wrapDocument(body));
zip.file("word/_rels/document.xml.rels", buildDocumentRels(extras));
zip.file(
	"docProps/core.xml",
	buildCoreProps({
		title: "Test Doc",
		author: "Kirill",
		created: "2026-05-01T00:00:00Z",
		modified: "2026-05-03T00:00:00Z",
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
