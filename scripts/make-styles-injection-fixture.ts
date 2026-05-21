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
	wrapDocument,
} from "./fixture-helpers";

/**
 * Build tests/fixtures/styles-injection.docx — a Word-canonical docx whose
 * body references paragraph and character styles that aren't defined in
 * styles.xml (which carries only Normal). Used to exercise the lazy
 * style-injection path in core/styles.tsx, plus the "from-scratch" path
 * via Pkg.deletePart in tests.
 *
 * Without ensureStyle, Word/LibreOffice fall back to Normal for the
 * pStyle/rStyle references. The fixture is canonical so Word opens it
 * without the "unreadable content / recover?" prompt that surfaces when
 * required parts (settings/fontTable/theme/webSettings/app) are missing.
 */

const root = resolve(import.meta.dir, "..");
const out = resolve(root, "tests/fixtures/styles-injection.docx");

const body = `
	<w:p>
		<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
		<w:r><w:t>Heading without a definition</w:t></w:r>
	</w:p>
	<w:p>
		<w:pPr><w:pStyle w:val="Quote"/></w:pPr>
		<w:r><w:t>A quoted paragraph that should render italic and indented.</w:t></w:r>
	</w:p>
	<w:p>
		<w:r><w:t xml:space="preserve">Inline </w:t></w:r>
		<w:r>
			<w:rPr><w:rStyle w:val="Code"/></w:rPr>
			<w:t>monospace</w:t>
		</w:r>
		<w:r><w:t xml:space="preserve"> reference.</w:t></w:r>
	</w:p>
	<w:p>
		<w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr>
		<w:r><w:t>Indented list-like item.</w:t></w:r>
	</w:p>
	${DEFAULT_SECTPR}`;

const zip = new JSZip();
zip.file("[Content_Types].xml", buildContentTypes());
zip.file("_rels/.rels", buildRootRels());
zip.file("word/document.xml", wrapDocument(body));
zip.file("word/_rels/document.xml.rels", buildDocumentRels());
zip.file(
	"docProps/core.xml",
	buildCoreProps({
		title: "Styles Injection Fixture",
		author: "docx-cli",
		created: "2026-05-21T00:00:00Z",
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
