import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import JSZip from "jszip";

// Pin per-entry ZIP mtimes to a fixed value so rebuilds are byte-deterministic.
// (`buildCoreProps` already takes fixed `created`/`modified` for `core.xml`;
// this addresses the per-entry mtime that JSZip would otherwise default to
// `new Date()`.)
process.env.DOCX_CLI_NOW ??= "2026-05-22T00:00:00Z";

import {
	addCanonicalParts,
	buildContentTypes,
	buildCoreProps,
	buildDocumentRels,
	buildRootRels,
	pinFixtureZipDates,
	wrapDocument,
} from "./helpers";

/**
 * Build tests/fixtures/transparent-wrappers.docx — a minimal docx exercising
 * <w:fldSimple> (cached field results) and <w:smartTag> (semantic
 * annotations). Both wrap runs that contribute to paragraph text but were
 * historically invisible to the XML-side offset arithmetic in
 * comments/helpers.tsx, replace/replace-span.tsx, and hyperlinks/wrap.tsx.
 *
 * Also an inline <w:sdt> content-control pair (#58): a Word form's checkbox
 * (`<w14:checkbox>`, glyph in `<w:sdtContent>`) and a Plain Text field — a
 * read-only shape we don't author, so it stays hand-rolled.
 *
 * Layout:
 *   p0: "Today is " + <w:fldSimple w:instr="DATE">"2026-05-05"</w:fldSimple> + "."
 *   p1: "Hello " + <w:smartTag w:element="PersonName">"Alice"</w:smartTag> + "."
 *   p2: <w:sdt checkbox>"☐"</w:sdt> + " Approved by " + <w:sdt alias="Approver">"Jane Doe"</w:sdt>
 */

const root = resolve(import.meta.dir, "../../..");
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
	<w:p>
		<w:sdt>
			<w:sdtPr>
				<w:id w:val="1001"/>
				<w14:checkbox>
					<w14:checked w14:val="0"/>
					<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>
					<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>
				</w14:checkbox>
			</w:sdtPr>
			<w:sdtContent>
				<w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic"/></w:rPr><w:t>☐</w:t></w:r>
			</w:sdtContent>
		</w:sdt>
		<w:r><w:t xml:space="preserve"> Approved by </w:t></w:r>
		<w:sdt>
			<w:sdtPr>
				<w:alias w:val="Approver"/>
				<w:id w:val="1002"/>
				<w:text/>
			</w:sdtPr>
			<w:sdtContent>
				<w:r><w:t>Jane Doe</w:t></w:r>
			</w:sdtContent>
		</w:sdt>
	</w:p>
	<w:sectPr/>`;

const zip = new JSZip();
pinFixtureZipDates(zip);
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
