import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import JSZip from "jszip";
import {
	addCanonicalParts,
	buildContentTypes,
	buildCoreProps,
	buildDocumentRels,
	buildRootRels,
	wrapDocument,
} from "./fixture-helpers";

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

const body = `
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
	<w:sectPr/>`;

const zip = new JSZip();
zip.file("[Content_Types].xml", buildContentTypes());
zip.file("_rels/.rels", buildRootRels());
zip.file("word/document.xml", wrapDocument(body));
zip.file("word/_rels/document.xml.rels", buildDocumentRels());
zip.file(
	"docProps/core.xml",
	buildCoreProps({
		title: "Transparent Wrappers Fixture",
		author: "docx-cli",
		created: "2026-05-05T00:00:00Z",
	}),
);
addCanonicalParts(zip);

const buf = await zip.generateAsync({
	type: "uint8array",
	compression: "DEFLATE",
});
mkdirSync(dirname(out), { recursive: true });
await Bun.write(out, buf);
console.log(`Wrote ${out} (${buf.length} bytes)`);
