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

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/tracked-moves.docx");

const body = `
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
		title: "Tracked Moves Fixture",
		author: "docx-cli",
		created: "2026-05-05T00:00:00Z",
		modified: "2026-05-05T12:00:00Z",
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
