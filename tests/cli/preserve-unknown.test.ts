import { describe, expect, test } from "bun:test";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import JSZip from "jszip";
import { runCli } from "./harness";

/**
 * Pillar invariant tests: any element we don't actively model survives every
 * mutating command unchanged. We use real OOXML elements that the AST doesn't
 * surface (`<w:bookmarkStart>`, `<w:bookmarkEnd>`, `<w:proofErr>`,
 * `<w:permStart>`) — they're zero-width paragraph-level markers that would
 * be silently destroyed if any walker stopped preserving unknown children.
 */

async function buildFixture(bodyXml: string, label: string): Promise<string> {
	const workspace = mkdtempSync(join(tmpdir(), `docx-cli-${label}-`));
	const docPath = join(workspace, "out.docx");

	const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
	<w:body>${bodyXml}<w:sectPr/></w:body>
</w:document>`;
	const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
	<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
	<Default Extension="xml" ContentType="application/xml"/>
	<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
	const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
	const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;

	const zip = new JSZip();
	zip.file("[Content_Types].xml", contentTypes);
	zip.file("_rels/.rels", rootRels);
	zip.file("word/document.xml", documentXml);
	zip.file("word/_rels/document.xml.rels", docRels);
	const buf = await zip.generateAsync({
		type: "uint8array",
		compression: "DEFLATE",
	});
	await Bun.write(docPath, buf);
	return docPath;
}

async function readDocumentXml(docPath: string): Promise<string> {
	const pkg = await Pkg.open(docPath);
	return pkg.readText("word/document.xml");
}

describe("preserve unknown elements — invariant tests", () => {
	test("replace preserves <w:bookmarkStart>/<w:bookmarkEnd> sitting between runs", async () => {
		const docPath = await buildFixture(
			`<w:p>` +
				`<w:r><w:t xml:space="preserve">before </w:t></w:r>` +
				`<w:bookmarkStart w:id="0" w:name="myBookmark"/>` +
				`<w:r><w:t>target</w:t></w:r>` +
				`<w:bookmarkEnd w:id="0"/>` +
				`<w:r><w:t xml:space="preserve"> after</w:t></w:r>` +
				`</w:p>`,
			"preserve-bookmark",
		);

		const result = await runCli("replace", docPath, "target", "REPLACED");
		expect(result.exitCode).toBe(0);

		const xml = await readDocumentXml(docPath);
		expect(xml).toContain('<w:bookmarkStart w:id="0" w:name="myBookmark"/>');
		expect(xml).toContain('<w:bookmarkEnd w:id="0"/>');
		expect(xml).toContain("REPLACED");
	});

	test("replace preserves <w:proofErr> markers when cutting nearby text", async () => {
		const docPath = await buildFixture(
			`<w:p>` +
				`<w:r><w:t xml:space="preserve">spell </w:t></w:r>` +
				`<w:proofErr w:type="spellStart"/>` +
				`<w:r><w:t>tihs</w:t></w:r>` +
				`<w:proofErr w:type="spellEnd"/>` +
				`<w:r><w:t xml:space="preserve"> word</w:t></w:r>` +
				`</w:p>`,
			"preserve-proof",
		);

		await runCli("replace", docPath, "tihs", "this");

		const xml = await readDocumentXml(docPath);
		expect(xml).toContain('<w:proofErr w:type="spellStart"/>');
		expect(xml).toContain('<w:proofErr w:type="spellEnd"/>');
		expect(xml).toContain("this");
	});

	test("replace preserves <w:permStart>/<w:permEnd> permission markers", async () => {
		const docPath = await buildFixture(
			`<w:p>` +
				`<w:permStart w:id="100" w:edGrp="everyone"/>` +
				`<w:r><w:t>locked content</w:t></w:r>` +
				`<w:permEnd w:id="100"/>` +
				`</w:p>`,
			"preserve-perm",
		);

		await runCli("replace", docPath, "locked content", "modified content");

		const xml = await readDocumentXml(docPath);
		expect(xml).toContain('<w:permStart w:id="100" w:edGrp="everyone"/>');
		expect(xml).toContain('<w:permEnd w:id="100"/>');
		expect(xml).toContain("modified content");
	});

	test("comments add preserves unknown markers in the same paragraph", async () => {
		const docPath = await buildFixture(
			`<w:p>` +
				`<w:r><w:t xml:space="preserve">Hello </w:t></w:r>` +
				`<w:bookmarkStart w:id="0" w:name="loc"/>` +
				`<w:r><w:t>world</w:t></w:r>` +
				`<w:bookmarkEnd w:id="0"/>` +
				`</w:p>`,
			"preserve-comment",
		);

		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--range",
			"p0:6-11",
			"--text",
			"about the world",
			"--author",
			"Tester",
		);
		expect(result.exitCode).toBe(0);

		const xml = await readDocumentXml(docPath);
		expect(xml).toContain('<w:bookmarkStart w:id="0" w:name="loc"/>');
		expect(xml).toContain('<w:bookmarkEnd w:id="0"/>');
		expect(xml).toContain("<w:commentRangeStart");
	});

	test("hyperlinks add preserves an unknown sibling marker", async () => {
		const docPath = await buildFixture(
			`<w:p>` +
				`<w:bookmarkStart w:id="0" w:name="anchor"/>` +
				`<w:r><w:t>click here</w:t></w:r>` +
				`<w:bookmarkEnd w:id="0"/>` +
				`</w:p>`,
			"preserve-hyperlink",
		);

		const result = await runCli(
			"hyperlinks",
			"add",
			docPath,
			"--at",
			"p0:0-10",
			"--url",
			"https://example.com",
		);
		expect(result.exitCode).toBe(0);

		const xml = await readDocumentXml(docPath);
		expect(xml).toContain('<w:bookmarkStart w:id="0" w:name="anchor"/>');
		expect(xml).toContain('<w:bookmarkEnd w:id="0"/>');
		expect(xml).toContain("<w:hyperlink");
	});

	test("non-checkbox SDT bodies are preserved AND their internal tracked changes aren't enumerated", async () => {
		// Word emits `<w:sdt>` for non-checkbox content controls too — Plain
		// Text, Rich Text, Dropdown, etc. — and a tracked edit inside one
		// shows up as `<w:ins>`/`<w:del>` inside `<w:sdtContent>`. The AST
		// reader and the apply walker MUST agree on whether to descend; if
		// they disagree, `tcN` ids drift between `track-changes list` and
		// `accept --at tcN`. Today neither walker descends, so the inner
		// revision is invisible to track-changes commands AND survives the
		// underlying XmlNode tree untouched.
		const docPath = await buildFixture(
			`<w:p>` +
				`<w:sdt>` +
				`<w:sdtPr>` +
				`<w:id w:val="42"/>` +
				`<w:placeholder><w:docPart w:val="DefaultPlaceholder_-1854013440"/></w:placeholder>` +
				`</w:sdtPr>` +
				`<w:sdtContent>` +
				`<w:ins w:id="7" w:author="Tester" w:date="2026-05-25T00:00:00Z">` +
				`<w:r><w:t>secret inside content control</w:t></w:r>` +
				`</w:ins>` +
				`</w:sdtContent>` +
				`</w:sdt>` +
				`</w:p>`,
			"sdt-tracked",
		);

		const list = await runCli("track-changes", "list", docPath);
		const changes = list.parsed as Array<{ id: string }>;
		// The inner `<w:ins>` MUST NOT surface — neither walker descends into
		// non-checkbox SDTs, so `track-changes list` shows zero entries.
		expect(changes).toEqual([]);

		// Round-trip through a save (via `replace` on text in a sibling, but
		// there is none here — use `insert` after p0 instead, which is enough
		// to force `Document.save`).
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"unrelated paragraph",
		);

		const xml = await readDocumentXml(docPath);
		// The whole SDT (including the inner ins) survives intact.
		expect(xml).toContain("<w:sdt>");
		expect(xml).toContain('<w:id w:val="42"/>');
		expect(xml).toContain(
			'<w:ins w:id="7" w:author="Tester" w:date="2026-05-25T00:00:00Z">',
		);
		expect(xml).toContain("secret inside content control");
	});

	test("find returns offsets that ignore unknown markers (markers don't shift indexing)", async () => {
		const docPath = await buildFixture(
			`<w:p>` +
				`<w:r><w:t xml:space="preserve">Hi </w:t></w:r>` +
				`<w:bookmarkStart w:id="0" w:name="x"/>` +
				`<w:r><w:t>there</w:t></w:r>` +
				`<w:bookmarkEnd w:id="0"/>` +
				`</w:p>`,
			"find-unknown",
		);

		const result = await runCli("find", docPath, "there");
		const payload = result.parsed as {
			matches: Array<{ locator: string }>;
		};
		// "Hi " is 3 chars; bookmarkStart is zero-width and unknown so it
		// contributes nothing to offset; "there" starts at 3.
		expect(payload.matches[0]?.locator).toBe("p0:3-8");
	});
});
