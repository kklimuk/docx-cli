import { resolve } from "node:path";
import { $ } from "bun";

// Pin core.xml timestamps so rebuilds are byte-deterministic.
process.env.DOCX_CLI_NOW ??= "2026-05-22T00:00:00Z";

/**
 * Build tests/fixtures/raw-ooxml.docx — dogfoods the `docx raw` escape hatch
 * end to end, exercising exactly the constructs the modeled verbs can't author:
 *
 *   - `raw insert` of a DROP-CAP paragraph (`<w:pPr><w:framePr w:dropCap>`);
 *   - `raw insert` of a table carrying a NESTED table inside its first cell;
 *   - `raw get s0` → `raw replace --at s0` (the section patch loop) adding
 *     line numbering (`<w:lnNumType>`), which no `sections` flag models;
 *   - the rel-then-body loop: `raw insert` of a `<Relationship>` (minting an
 *     rId in word/_rels/document.xml.rels), then a body `<w:hyperlink r:id>`
 *     referencing it — the workflow OLE/chart/external-link fragments need;
 *   - `raw part add` of a binary embeddings part + a relationship targeting
 *     it (the internal-Target gate resolving against a raw-added part);
 *   - `raw edit` patching the drop cap's `w:lines` in place — the
 *     get → modify → replace loop as one gated call.
 *
 * Building it is itself a smoke test for the gate pipeline (every fragment
 * passes well-formedness, addressable roots, child order, namespace, reference
 * and full-document XSD gates), and the LibreOffice round-trip (CORE_FIXTURES)
 * proves the spliced XML is render-valid — the axis the unit tests can't see.
 *
 * Coverage note (deliberate): `raw` has NO weak-agent scenario, by decision —
 * it is an expert surface for wrapper-builders, framed as last-resort in help;
 * weak agents should keep reaching for the modeled verbs.
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/raw-ooxml.docx");
const cliEntry = resolve(root, "src/index.ts");

async function cli(...args: string[]): Promise<string> {
	const result = await $`bun ${cliEntry} ${args}`.quiet();
	return result.stdout.toString();
}

await cli(
	"create",
	out,
	"--force",
	"--text",
	"riginal opening paragraph, styled around a drop cap.",
);

// A drop cap the `insert` verb has no flag for: replace the plain opener with
// a framePr drop-cap "O" paragraph ahead of it.
await cli(
	"raw",
	"insert",
	out,
	"--before",
	"p0",
	"--xml",
	'<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="2" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/><w:rPr><w:sz w:val="72"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="72"/></w:rPr><w:t>O</w:t></w:r></w:p>',
);

// Deepen the drop cap in place — dogfoods `raw edit` (find/with patch).
await cli(
	"raw",
	"edit",
	out,
	"--at",
	"p0",
	"--find",
	'w:lines="2"',
	"--with",
	'w:lines="3"',
);

// A nested table (a table inside a cell) — `tables create` models flat tables.
await cli(
	"raw",
	"insert",
	out,
	"--at-end",
	"--xml",
	[
		'<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>',
		'<w:tblGrid><w:gridCol w:w="4675"/><w:gridCol w:w="4675"/></w:tblGrid>',
		"<w:tr><w:tc><w:p><w:r><w:t>outer cell with nested table</w:t></w:r></w:p>",
		'<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
		"<w:p/></w:tc>",
		"<w:tc><w:p><w:r><w:t>plain cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
	].join(""),
);

// The rel-then-body loop: mint a hyperlink relationship raw, then splice a
// paragraph whose <w:hyperlink r:id> references it — proving the minted rId
// satisfies the dangling-rId gate and survives the LibreOffice round-trip.
const relOutput = await cli(
	"raw",
	"insert",
	out,
	"--xml",
	'<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/raw" TargetMode="External"/>',
);
const rId = relOutput.trim().split("\n")[0];
if (!rId?.match(/^rId\d+$/))
	throw new Error(`expected a minted rId, got: ${relOutput}`);
await cli(
	"raw",
	"insert",
	out,
	"--at-end",
	"--xml",
	`<w:p><w:hyperlink r:id="${rId}"><w:r><w:t>raw-minted link</w:t></w:r></w:hyperlink></w:p>`,
);

// The part loop: create a binary part raw, then a relationship targeting it —
// an unreferenced part + rel pair is harmless by the relationships invariant,
// and building it proves the part gates + the internal-Target resolution.
const blobPath = resolve(root, "tests/fixtures/setup/.raw-ooxml-blob.tmp");
await Bun.write(blobPath, new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0x00, 0x01]));
await cli(
	"raw",
	"part",
	"add",
	out,
	"--name",
	"word/embeddings/object1.bin",
	"--from",
	blobPath,
	"--content-type",
	"application/vnd.openxmlformats-officedocument.oleObject",
);
await Bun.file(blobPath).delete();
await cli(
	"raw",
	"insert",
	out,
	"--xml",
	'<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/object1.bin"/>',
);

// The section patch loop: get the exact sectPr, add line numbering, put it back.
const sectPr = (await cli("raw", "get", out, "--at", "s0")).trim();
const numbered = sectPr.replace(
	"<w:cols",
	'<w:lnNumType w:countBy="1" w:restart="continuous"/><w:cols',
);
if (numbered === sectPr) throw new Error("expected <w:cols> in the sectPr");
await cli("raw", "replace", out, "--at", "s0", "--xml", numbered);

console.log(`built ${out}`);
