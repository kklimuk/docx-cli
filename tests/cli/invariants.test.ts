import { describe, expect, test } from "bun:test";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import type { XmlNode } from "@core/parser";
import JSZip from "jszip";
import { runCli, tempWorkspace } from "./harness";
import { freshFixture, readDocumentXml } from "./helpers";

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
			"--at",
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

	test("replace preserves an UNMODELED run property (<w:emboss>) in the rPr", async () => {
		// We now model <w:shd> and <w:color w:themeColor> etc., but a run property
		// we don't model (e.g. <w:emboss>) must still survive a mutation untouched
		// — the AST is a view; existing runs are mutated in place, never re-emitted
		// through the RunProperties emitter.
		const docPath = await buildFixture(
			`<w:p><w:r><w:rPr><w:emboss/></w:rPr><w:t>embossed</w:t></w:r></w:p>`,
			"preserve-emboss",
		);

		await runCli("replace", docPath, "embossed", "EMBOSSED");

		const xml = await readDocumentXml(docPath);
		expect(xml).toContain("<w:emboss/>");
		expect(xml).toContain("EMBOSSED");
	});
});

// edit and delete are the heaviest mutators (run-splitting / LCS surgery and
// whole-block removal), so they are the likeliest to drop adjacent unmodeled
// markers. The replace/comments/hyperlinks paths are covered above; these lock
// the same invariant for edit (whole-paragraph + span) and delete (neighbor).
describe("edit / delete preserve unknown elements", () => {
	const markedParagraphs =
		`<w:p>` +
		`<w:r><w:t xml:space="preserve">before </w:t></w:r>` +
		`<w:bookmarkStart w:id="0" w:name="bm"/>` +
		`<w:proofErr w:type="spellStart"/>` +
		`<w:r><w:t>target</w:t></w:r>` +
		`<w:proofErr w:type="spellEnd"/>` +
		`<w:bookmarkEnd w:id="0"/>` +
		`<w:r><w:t xml:space="preserve"> after</w:t></w:r>` +
		`</w:p>` +
		`<w:p><w:r><w:t>second paragraph</w:t></w:r></w:p>`;

	test("whole-paragraph edit keeps bookmark + proofErr siblings", async () => {
		const docPath = await buildFixture(markedParagraphs, "edit-whole-unknown");

		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--text",
			"rewritten",
		);
		expect(result.exitCode).toBe(0);

		const xml = await readDocumentXml(docPath);
		expect(xml).toContain('<w:bookmarkStart w:id="0" w:name="bm"/>');
		expect(xml).toContain('<w:bookmarkEnd w:id="0"/>');
		expect(xml).toContain('<w:proofErr w:type="spellStart"/>');
		expect(xml).toContain("rewritten");
	});

	test("span edit keeps bookmark + proofErr siblings", async () => {
		const docPath = await buildFixture(markedParagraphs, "edit-span-unknown");

		// Span the leading "before " run only; the markers sit just past it.
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p0:0-6",
			"--text",
			"BEFORE",
		);
		expect(result.exitCode).toBe(0);

		const xml = await readDocumentXml(docPath);
		expect(xml).toContain('<w:bookmarkStart w:id="0" w:name="bm"/>');
		expect(xml).toContain('<w:proofErr w:type="spellEnd"/>');
	});

	test("deleting a neighbor paragraph leaves the marked paragraph intact", async () => {
		const docPath = await buildFixture(
			markedParagraphs,
			"delete-neighbor-unknown",
		);

		const result = await runCli("delete", docPath, "--at", "p1");
		expect(result.exitCode).toBe(0);

		const xml = await readDocumentXml(docPath);
		expect(xml).toContain('<w:bookmarkStart w:id="0" w:name="bm"/>');
		expect(xml).toContain('<w:bookmarkEnd w:id="0"/>');
		expect(xml).not.toContain("second paragraph");
	});
});

/**
 * Regression coverage for the AST/XML offset alignment fix: when a paragraph
 * contains a transparent wrapper (<w:fldSimple>, <w:smartTag>) the AST
 * surfaces the inner text as part of `paragraph.runs`, so `find` reports
 * offsets that include the wrapper's content. Before the fix, the XML-side
 * traversal in `comments/helpers.tsx`, `replace/replace-span.tsx`, and
 * `hyperlinks/wrap.tsx` skipped these wrappers, producing offsets that
 * disagreed with the AST and breaking find→replace / find→comments-add
 * pipelines.
 *
 * Fixture (tests/fixtures/transparent-wrappers.docx):
 *   p0: "Today is " + <w:fldSimple>"2026-05-05"</w:fldSimple> + "."
 *   p1: "Hello " + <w:smartTag>"Alice"</w:smartTag> + "."
 */
const FIXTURE = "tests/fixtures/transparent-wrappers.docx";

const freshCopy = (label: string) => freshFixture(label, FIXTURE);

describe("transparent wrappers — find offsets include inner text", () => {
	test("find locates text inside <w:fldSimple>", async () => {
		const result = await runCli("find", FIXTURE, "2026-05-05");
		const payload = result.parsed as {
			totalMatches: number;
			matches: Array<{ locator: string; text: string }>;
		};
		expect(payload.totalMatches).toBe(1);
		expect(payload.matches[0]?.text).toBe("2026-05-05");
		// "Today is " is 9 chars; the field content starts at offset 9.
		expect(payload.matches[0]?.locator).toBe("p0:9-19");
	});

	test("find locates text inside <w:smartTag>", async () => {
		const result = await runCli("find", FIXTURE, "Alice");
		const payload = result.parsed as {
			totalMatches: number;
			matches: Array<{ locator: string; text: string }>;
		};
		expect(payload.totalMatches).toBe(1);
		// "Hello " is 6 chars; "Alice" starts at offset 6.
		expect(payload.matches[0]?.locator).toBe("p1:6-11");
	});

	test("find locates text spanning past the wrapper", async () => {
		// "is 2026" crosses the boundary into <w:fldSimple>.
		const result = await runCli("find", FIXTURE, "is 2026");
		const payload = result.parsed as {
			totalMatches: number;
			matches: Array<{ locator: string }>;
		};
		expect(payload.totalMatches).toBe(1);
		expect(payload.matches[0]?.locator).toBe("p0:6-13");
	});
});

describe("transparent wrappers — replace inside the wrapper", () => {
	test("replace text fully inside <w:fldSimple> swaps the cached field result", async () => {
		const docPath = await freshCopy("transparent-replace-fld");
		const result = await runCli("replace", docPath, "2026-05-05", "today");
		expect(result.exitCode).toBe(0);

		// Re-read and confirm the new text appears at the right paragraph.
		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as {
			blocks: Array<{
				type: string;
				runs?: Array<{ type: string; text?: string }>;
			}>;
		};
		const p0Text = (doc.blocks[0]?.runs ?? [])
			.filter((run) => run.type === "text")
			.map((run) => run.text ?? "")
			.join("");
		expect(p0Text).toBe("Today is today.");

		// The "." run on the right side of the fldSimple should still be there.
		expect(p0Text.endsWith(".")).toBe(true);
	});

	test("replace text fully inside <w:smartTag> swaps the inner content", async () => {
		const docPath = await freshCopy("transparent-replace-smart");
		const result = await runCli("replace", docPath, "Alice", "Bob");
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as {
			blocks: Array<{
				type: string;
				runs?: Array<{ type: string; text?: string }>;
			}>;
		};
		const p1Text = (doc.blocks[1]?.runs ?? [])
			.filter((run) => run.type === "text")
			.map((run) => run.text ?? "")
			.join("");
		expect(p1Text).toBe("Hello Bob.");
	});

	test("replace span crossing INTO a wrapper splits it cleanly", async () => {
		const docPath = await freshCopy("transparent-replace-cross");
		// "is 2026" crosses the fldSimple boundary; replacing it should leave
		// the rest of the date inside a (now-shorter) fldSimple half.
		const result = await runCli("replace", docPath, "is 2026", "was 2025");
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as {
			blocks: Array<{
				type: string;
				runs?: Array<{ type: string; text?: string }>;
			}>;
		};
		const p0Text = (doc.blocks[0]?.runs ?? [])
			.filter((run) => run.type === "text")
			.map((run) => run.text ?? "")
			.join("");
		expect(p0Text).toBe("Today was 2025-05-05.");
	});

	test("the fldSimple wrapper survives a fully-internal replace", async () => {
		const docPath = await freshCopy("transparent-replace-survives");
		await runCli("replace", docPath, "2026-05-05", "today");

		const pkg = await Pkg.open(docPath);
		const xml = await pkg.readText("word/document.xml");
		// fldSimple still wraps the (replaced) inner content.
		expect(xml).toContain("<w:fldSimple");
	});
});

describe("transparent wrappers — comments add against the wrapper", () => {
	test("comments add accepts a span that ends past <w:fldSimple>", async () => {
		const docPath = await freshCopy("transparent-comment-fld");
		// p0 length is 20: "Today is " (9) + "2026-05-05" (10) + "." (1).
		// Range covers "is 2026-05-05" (offsets 6-19).
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			"p0:6-19",
			"--text",
			"verify the date",
			"--author",
			"Tester",
		);
		expect(result.exitCode).toBe(0);
	});

	test("comments add accepts a span that ends past <w:smartTag>", async () => {
		const docPath = await freshCopy("transparent-comment-smart");
		// p1 length is 12: "Hello " (6) + "Alice" (5) + "." (1).
		// Range covers "Alice." (offsets 6-12).
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			"p1:6-12",
			"--text",
			"check the name",
			"--author",
			"Tester",
		);
		expect(result.exitCode).toBe(0);
	});

	test("comments add against a span fully past the wrapper succeeds with correct length", async () => {
		const docPath = await freshCopy("transparent-comment-past");
		// p0 length is 20; targeting the trailing "." at offset 19-20 should work.
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			"p0:19-20",
			"--text",
			"period",
			"--author",
			"Tester",
		);
		expect(result.exitCode).toBe(0);
	});
});

/**
 * Regression coverage for Bug A (parser dropped <w:t> siblings of inline
 * children) and Bug B (sliceRun cloned non-text inline markers into every
 * slice). The mnda.docx fixture has two <w:r> elements that put a <w:tab/>
 * and a <w:t> in the same run:
 *   <w:r>...<w:tab/><w:t xml:space="preserve">Expires </w:t></w:r>
 *   <w:r>...<w:tab/><w:t>Continues until terminated ...</w:t></w:r>
 */
describe("mnda fixture — tab + text in same <w:r>", () => {
	const FIXTURE = "tests/fixtures/mnda.docx";

	test("Bug A: AST exposes <w:t> text that sits after a <w:tab/> sibling", async () => {
		const result = await runCli("read", FIXTURE, "--ast");
		const allText = collectAllText(result.parsed);
		expect(allText).toContain("Expires ");
		expect(allText).toContain(
			"Continues until terminated in accordance with the terms of the MNDA.",
		);
	});

	test("Bug A: find locates text after a tab in the same run", async () => {
		const result = await runCli("find", FIXTURE, "Continues until terminated");
		const payload = result.parsed as {
			totalMatches: number;
			matches: Array<{ text: string }>;
		};
		expect(payload.totalMatches).toBeGreaterThanOrEqual(1);
		expect(payload.matches[0]?.text).toBe("Continues until terminated");
	});

	test("Bug B: track-changes-on replace of '1 year(s)' wraps only the year run; the adjacent tab+Expires run is untouched", async () => {
		const workspace = tempWorkspace("mnda-replace");
		const docPath = join(workspace, "mnda.docx");
		await Bun.write(docPath, Bun.file(FIXTURE));

		await runCli("track-changes", docPath, "on");
		const replaceResult = await runCli(
			"replace",
			docPath,
			"1 year(s)",
			"2 year(s)",
			"--all",
			"--author",
			"Tester",
		);
		expect(replaceResult.exitCode).toBe(0);

		const pkg = await Pkg.open(docPath);
		const documentXml = await pkg.readText("word/document.xml");

		// The "Expires " run should still be a single <w:r> carrying the tab and
		// the original text, not split or duplicated by sliceRun.
		const expiresRuns = matchAll(
			documentXml,
			/<w:r\b[^>]*>(?:(?!<\/w:r>).)*?<w:tab\/>(?:(?!<\/w:r>).)*?<w:t[^>]*>Expires <\/w:t>(?:(?!<\/w:r>).)*?<\/w:r>/g,
		);
		expect(expiresRuns).toHaveLength(1);

		// The "Continues until terminated" run is the same shape — verify it
		// also survives untouched (the replace span doesn't reach this paragraph,
		// but XmlNode.serialize round-tripping should not damage it either).
		const continuesRuns = matchAll(
			documentXml,
			/<w:r\b[^>]*>(?:(?!<\/w:r>).)*?<w:tab\/>(?:(?!<\/w:r>).)*?<w:t[^>]*>Continues until terminated[^<]*<\/w:t>(?:(?!<\/w:r>).)*?<\/w:r>/g,
		);
		expect(continuesRuns).toHaveLength(1);

		// "1 year(s)" appears twice in the source. Each match should produce one
		// <w:del> wrapping the original run and one adjacent <w:ins> carrying
		// "2 year(s)".
		const delsWith1Year = matchAll(
			documentXml,
			/<w:del\b[^>]*w:author="Tester"[^>]*>(?:(?!<\/w:del>).)*?<w:delText[^>]*>1 year\(s\)<\/w:delText>(?:(?!<\/w:del>).)*?<\/w:del>/g,
		);
		const insWith2Year = matchAll(
			documentXml,
			/<w:ins\b[^>]*w:author="Tester"[^>]*>(?:(?!<\/w:ins>).)*?<w:t[^>]*>2 year\(s\)<\/w:t>(?:(?!<\/w:ins>).)*?<\/w:ins>/g,
		);
		expect(delsWith1Year).toHaveLength(2);
		expect(insWith2Year).toHaveLength(2);

		// And the original literal "1 year(s)" should no longer appear inside
		// any <w:t> (only inside <w:delText> within <w:del>).
		const liveOneYear = matchAll(documentXml, /<w:t[^>]*>1 year\(s\)<\/w:t>/g);
		expect(liveOneYear).toHaveLength(0);
	});
});

function collectAllText(parsed: unknown): string {
	const doc = parsed as {
		blocks: Array<{
			type: string;
			runs?: Array<{ type: string; text?: string }>;
			rows?: Array<{
				cells: Array<{
					blocks: Array<{
						type: string;
						runs?: Array<{ type: string; text?: string }>;
					}>;
				}>;
			}>;
		}>;
	};
	const out: string[] = [];
	function walk(
		blocks: Array<{
			type: string;
			runs?: Array<{ type: string; text?: string }>;
			rows?: Array<{
				cells: Array<{
					blocks: Array<{
						type: string;
						runs?: Array<{ type: string; text?: string }>;
					}>;
				}>;
			}>;
		}>,
	): void {
		for (const block of blocks) {
			if (block.runs) {
				for (const run of block.runs) {
					if (run.type === "text" && typeof run.text === "string") {
						out.push(run.text);
					}
				}
			}
			if (block.rows) {
				for (const row of block.rows) {
					for (const cell of row.cells) {
						walk(cell.blocks);
					}
				}
			}
		}
	}
	walk(doc.blocks);
	return out.join("");
}

function matchAll(haystack: string, pattern: RegExp): string[] {
	const out: string[] = [];
	for (const match of haystack.matchAll(pattern)) {
		out.push(match[0]);
	}
	return out;
}

// Catches the bug class that produced Word's "unreadable content" dialog: a part
// that uses an XML namespace prefix (`r:id` on a note-body `<w:hyperlink>`) the
// part never declares — malformed XML — and the related "dangling rId" class.
// Our LibreOffice-render + AST-roundtrip tests missed both because LibreOffice
// is permissive and the AST never re-parses namespaces. These checks are the
// guard. (`xmllint` is run too when present — it independently flags undeclared
// prefixes — but the structured checks below always run in CI.)

const XMLLINT = Bun.which("xmllint");
const PNG_PATH = join(
	import.meta.dir,
	"..",
	"fixtures",
	"assets",
	"sample.png",
);

describe("generated .docx is valid OOXML", () => {
	test("markdown with a reused footnote containing a hyperlink (the regression)", async () => {
		const docPath = join(tempWorkspace("validity-fn-link"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"First cite[^a] then again[^a].\n\n[^a]: see [AP](https://ap.example.com/x) and [NBC](https://nbc.example.com/y).",
		);
		await assertValidDocx(docPath);
	});

	test("body hyperlink + heading + footnote-with-link in one doc", async () => {
		const docPath = join(tempWorkspace("validity-mixed"), "out.docx");
		await runCli("create", docPath, "--text", "seed");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"# Title\n\nA [body link](https://example.com/a) and a note[^n].\n\n[^n]: with [a link](https://example.com/b).",
		);
		await assertValidDocx(docPath);
	});

	test("hyperlinks add (body relationship)", async () => {
		const docPath = join(tempWorkspace("validity-hyperlink"), "out.docx");
		await runCli("create", docPath, "--text", "Click this phrase here.");
		await runCli(
			"hyperlinks",
			"add",
			docPath,
			"--at",
			"p0:6-17",
			"--url",
			"https://example.com",
		);
		await assertValidDocx(docPath);
	});

	test("footnotes add (no links — the plain note part stays valid)", async () => {
		const docPath = join(tempWorkspace("validity-fn-add"), "out.docx");
		await runCli("create", docPath, "--text", "A paragraph.");
		await runCli("footnotes", "add", docPath, "--at", "p0", "--text", "a note");
		await assertValidDocx(docPath);
	});

	// The note-body markdown path (`--markdown` on add/edit) mints hyperlink rels
	// and must route them into the NOTE part's own rels — the same class of bug
	// the regression above guards, but reached through a different command.
	test("footnotes add --markdown with a body hyperlink (note-part rels)", async () => {
		const docPath = join(tempWorkspace("validity-fn-add-md"), "out.docx");
		await runCli("create", docPath, "--text", "A paragraph.");
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--markdown",
			"See [AP](https://ap.example.com/x) for the long form.",
		);
		await assertValidDocx(docPath);
	});

	test("endnotes edit --markdown with a body hyperlink", async () => {
		const docPath = join(tempWorkspace("validity-en-edit-md"), "out.docx");
		await runCli("create", docPath, "--text", "A paragraph.");
		await runCli(
			"endnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"placeholder",
		);
		await runCli(
			"endnotes",
			"edit",
			docPath,
			"--at",
			"en1",
			"--markdown",
			"Revised with [a source](https://example.com/src).",
		);
		await assertValidDocx(docPath);
	});

	// A markdown image in a note body is dropped (note bodies are text + links);
	// if the strip ever regresses the image embeds with its media rel in
	// document.xml.rels while the <w:drawing r:embed> lands in footnotes.xml — a
	// dangling rId the relationship check below catches.
	test("footnotes add --markdown with an image strips it (no dangling media rel)", async () => {
		const docPath = join(tempWorkspace("validity-fn-add-img"), "out.docx");
		await runCli("create", docPath, "--text", "A paragraph.");
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--markdown",
			`Caption text ![pic](${PNG_PATH}) and more.`,
		);
		await assertValidDocx(docPath);
	});
});

/** Assert every XML part of `docPath` is namespace-well-formed and every
 *  relationship reference resolves. */
async function assertValidDocx(docPath: string): Promise<void> {
	const pkg = await Pkg.open(docPath);
	const xmlParts = pkg
		.listParts()
		.filter((name) => name.endsWith(".xml") || name.endsWith(".rels"));

	for (const name of xmlParts) {
		const text = await pkg.readText(name);

		// (1) Every namespace prefix used on a tag or attribute key is declared.
		// This is the exact invariant the footnote-hyperlink bug violated (`r:id`
		// used, `xmlns:r` never declared on `footnotes.xml`).
		const tree = await pkg.readPart(name);
		assertNamespacesDeclared(name, tree ?? []);

		// (2) Every `r:id` / `r:embed` / `r:link` resolves to a relationship in
		// THIS part's own rels (a dangling rId is the other corruption class).
		await assertRelationshipsResolve(pkg, name, tree ?? []);

		// (3) Belt-and-suspenders: libxml2's namespace-aware parser, when present,
		// independently rejects undeclared prefixes + any other malformedness.
		if (XMLLINT) await assertXmllintClean(name, text);
	}
}

function assertNamespacesDeclared(partName: string, tree: XmlNode[]): void {
	const declared = new Set<string>(["xml"]);
	const used = new Set<string>();
	const visit = (node: XmlNode): void => {
		const tagColon = node.tag.indexOf(":");
		if (tagColon > 0) used.add(node.tag.slice(0, tagColon));
		for (const key of Object.keys(node.attributes)) {
			if (key === "xmlns") continue;
			if (key.startsWith("xmlns:")) {
				declared.add(key.slice("xmlns:".length));
				continue;
			}
			const colon = key.indexOf(":");
			if (colon > 0) used.add(key.slice(0, colon));
		}
		for (const child of node.children) visit(child);
	};
	for (const root of tree) visit(root);

	for (const prefix of used) {
		expect(
			declared.has(prefix),
			`${partName}: namespace prefix "${prefix}:" is used but never declared (missing xmlns:${prefix}) — malformed XML, which Word reports as "unreadable content"`,
		).toBe(true);
	}
}

async function assertRelationshipsResolve(
	pkg: Pkg,
	partName: string,
	tree: XmlNode[],
): Promise<void> {
	const used = new Set<string>();
	const visit = (node: XmlNode): void => {
		for (const [key, value] of Object.entries(node.attributes)) {
			if (
				(key === "r:id" || key === "r:embed" || key === "r:link") &&
				typeof value === "string"
			) {
				used.add(value);
			}
		}
		for (const child of node.children) visit(child);
	};
	for (const root of tree) visit(root);
	if (used.size === 0) return;

	const relsName = relsPartNameFor(partName);
	const relsTree = await pkg.readPart(relsName);
	const declared = new Set<string>();
	for (const root of relsTree ?? []) {
		const visitRel = (node: XmlNode): void => {
			if (node.tag === "Relationship") {
				const id = node.attributes.Id;
				if (typeof id === "string") declared.add(id);
			}
			for (const child of node.children) visitRel(child);
		};
		visitRel(root);
	}

	for (const rId of used) {
		expect(
			declared.has(rId),
			`${partName}: references relationship "${rId}" but ${relsName} has no such <Relationship> (dangling rId — Word reports "unreadable content")`,
		).toBe(true);
	}
}

async function assertXmllintClean(
	partName: string,
	xml: string,
): Promise<void> {
	const proc = Bun.spawn(["xmllint", "--noout", "-"], {
		stdin: "pipe",
		stdout: "ignore",
		stderr: "pipe",
	});
	proc.stdin.write(xml);
	await proc.stdin.end();
	const [code, stderr] = await Promise.all([
		proc.exited,
		new Response(proc.stderr).text(),
	]);
	expect(code, `${partName} is not namespace-well-formed:\n${stderr}`).toBe(0);
}

function relsPartNameFor(partName: string): string {
	const slash = partName.lastIndexOf("/");
	const dir = partName.slice(0, slash);
	const base = partName.slice(slash + 1);
	return `${dir}/_rels/${base}.rels`;
}
